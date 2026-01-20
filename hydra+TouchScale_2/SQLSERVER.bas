Attribute VB_Name = "SQLSERVER"
Option Explicit

Public MyConnObj As New ADODB.Connection
Public MyRecSet As New ADODB.Recordset
Public strStream As New ADODB.Stream
Public bbDDver As Integer

Dim MyMinSec As Integer
#Const BALSAM = False

#If BALSAM = True Then
Public Type VitaMin
    daily_ref As Double
    name As String
    type As String
    units As String
End Type
Public Type D_VM
    data(26) As VitaMin
End Type
Public MyVM As D_VM

Public Type Item_Nutrition_Card
    Code As Long
    Card As Long
End Type

Public Type Nutrition_Base
    Card As Long
    portion_weight As Double
    energy_kj_100 As Double
    energy_kj_portion As Double
    energy_kcal_100 As Double
    energy_kcal_portion As Double
    fat_100 As Double
    fat_portion As Double
    saturates_100 As Double
    saturates_portion As Double
    mono_unsaturates_100 As Double
    mono_unsaturates_portion As Double
    polyunsaturates_100 As Double
    polyunsaturates_portion As Double
    carbohydrate_100 As Double
    carbohydrate_portion As Double
    sugars_100 As Double
    sugars_portion As Double
    polyols_100 As Double
    polyols_portion As Double
    starch_100 As Double
    starch_portion As Double
    fibre_100 As Double
    fibre_portion As Double
    protein_100 As Double
    protein_portion As Double
    salt_100 As Double
    salt_portion  As Double
End Type

Public Type Nutrition_Vita_Mine
    Card As Long
    identifier As Long '0 a 26 (abajo identificacion)
    'type As String 'V --> Vitamina, M --> Minerales
    'name As String 'abajo nombre en función identificacion
    'units As String 'abajo unidades en función identificacion
    'daily_ref As Double 'abajo referencia en funcion identificacion
    per100 As Double
    perportion As Double
End Type
'identifier type    name                units   daily_ref
'0          V       Vitamina A          ?g      800
'1          V       Vitamina D          ?g      5
'2          V       Vitamina E          mg      12
'3          V       Vitamina K          ?g      75
'4          V       Vitamina C          mg      80
'5          V       Tiamina             mg      1,1
'6          V       Riboflavina         mg      1,4
'7          V       Niacina             mg      16
'8          V       Vitamina B6         mg      1,4
'9          V       Ácido fólico        ?g      200
'10         V       Vitamina B12        ?g      2,5
'11         V       Biotina             ?g      50
'12         V       Ácido pantoténico   mg      6
'13         M       Potasio             mg      2000
'14         M       Cloro               mg      800
'15         M       Calcio              mg      800
'16         M       Fósforo             mg      700
'17         M       Magnesio            mg      375
'18         M       Hierro              mg      14
'19         M       Zinc                mg      10
'20         M       Cobre               mg      1
'21         M       Manganeso           mg      2
'22         M       Flúor               mg      3,5
'23         M       Selenio             ?g      55
'24         M       Cromo               ?g      40
'25         M       Molibdeno           ?g      50
'26         M       Yodo                ?g      150

Public Sub Absorver_Estructura()

    'Borrar datos previos existentes...
    'Revisar actual Hydra para ver netID
    CancelHumano = False
    canCelar = 0
    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = True
    frmMonitor.Label1.Caption = ""

    MostrarMonitor

    frmControl.Switch_mode 0
    If EstadoBal = 0 Then 'Protección mínima...
        frmMonitor.MostrarInfo CargaCadena(329) & " ..."  '"Secciones..."
        Call LeeSecciones
        'frmMonitor.MostrarInfo CargaCadena(303) & " ..."  '"Equipos..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeEquipos
        frmMonitor.MostrarInfo CargaCadena(67) & " ..." '"Familias..."
        If canCelar = 0 And CancelHumano = False Then Call LeeFamilias
        frmMonitor.MostrarInfo CargaCadena(13) & " ..." '"Vendedores..."
        If canCelar = 0 And CancelHumano = False Then Call LeeVendedores
        frmMonitor.MostrarInfo CargaCadena(214) & " ..." '"Ivas..."
        If canCelar = 0 And CancelHumano = False Then Call LeeIVAS
        frmMonitor.MostrarInfo CargaCadena(1084) & " ..." '"Articulos..."
        If canCelar = 0 And CancelHumano = False Then Call LeeArticulos
        'TouchScale ... prueba captura paises...
        frmMonitor.MostrarInfo CargaCadena(619) & " ..." '"Fichas de vacuno..."
        If canCelar = 0 And CancelHumano = False Then Call capturaPaises
        If canCelar = 0 And CancelHumano = False Then Call capturaFichas
        frmMonitor.MostrarInfo CargaCadena(23) & " ..." '"Códigos de barras..."
        If canCelar = 0 And CancelHumano = False Then Call capturaCB
        frmMonitor.MostrarInfo CargaCadena(1056) & " ..." '"Clientes..."
        If canCelar = 0 And CancelHumano = False Then Call capturaClientes
        ''''''''''''''''''''''''''''''''''''''''

        Call Corregir_Nulos

    End If
    'cerrar conexión
    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0
    EstadoBal = 1000
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.Command1.Enabled = True
    HabilitarSalir
    frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
    If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
End Sub

Public Function CodiIdent_FS(ByVal FS As Long) As Long
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim nFs As Long
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    Set myRec = Mybase.OpenRecordset("select * from seccion where secc_maqui=" & CStr(FS))
    If Not myRec.EOF Then
        nFs = myRec.Fields("codi_ident")
    Else
        myRec.Close
        Set myRec = Mybase.OpenRecordset("select * from seccion")
        If Not myRec.EOF Then
            myRec.MoveFirst
            nFs = myRec.Fields("codi_ident")
        End If
    End If
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    If Err.Number <> 0 Then
        CodiIdent_FS = 0
    Else
        CodiIdent_FS = nFs
    End If
    On Error GoTo 0
End Function

'DROP TABLE IF EXISTS `mbd`.`Items`;
'CREATE TABLE  `mbd`.`Items` (
'  `Id` int(11) DEFAULT NULL,
'  `AlfaCode` varchar(50) DEFAULT '',
'  `Name` varchar(100) DEFAULT '',
'  `Plu` int(11) DEFAULT '0',
'  `Family` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Price100g` int(11) DEFAULT '0',
'  `SaleForm` int(11) DEFAULT '0',
'  `VAT` int(11) DEFAULT '0',
'  `EAN13` varchar(13) DEFAULT '',
'  `EAN14` varchar(14) DEFAULT '',
'  `Discount` decimal(4,2) DEFAULT '0.00',
'  `Expiring` int(11) DEFAULT '0',
'  `BestBefore` int(11) DEFAULT '0',
'  `RemoveDate` int(11) DEFAULT '0',
'  `FreezingDate` datetime DEFAULT '0000-00-00 00:00:00',
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `Label` int(11) DEFAULT '0',
'  `Label2` int(11) DEFAULT '0',
'  `StepPrice` varchar(1) DEFAULT '0',
'  `Icon` varchar(100) DEFAULT '',
'  `Mixed` int(11) DEFAULT '0',
'  `MixedWeight` decimal(5,3) DEFAULT '0.000',
'  `MinWeight` decimal(5,3) DEFAULT '0.000',
'  `MaxWeight` decimal(5,3) DEFAULT '0.000',
'  `ApproxWeight` decimal(5,3) DEFAULT '0.000',
'  `Loss` decimal(5,3) DEFAULT '0.000',
'  `TotPresel` int(11) DEFAULT '0',
'  `IdCard` int(11) DEFAULT '0',
'  `IdBovine` int(11) DEFAULT '0',
'  `IdNutritional` int(11) DEFAULT '0',
'  `Text01` varchar(100) DEFAULT '',
'  `Text02` varchar(100) DEFAULT '',
'  `Text03` varchar(100) DEFAULT '',
'  `Text04` varchar(100) DEFAULT '',
'  `Text05` varchar(100) DEFAULT '',
'  `Text06` varchar(100) DEFAULT '',
'  `Text07` varchar(100) DEFAULT '',
'  `Text08` varchar(100) DEFAULT '',
'  `Text09` varchar(100) DEFAULT '',
'  `Text10` varchar(100) DEFAULT '',
'  `Text11` varchar(100) DEFAULT '',
'  `Text12` varchar(100) DEFAULT '',
'  `Text13` varchar(100) DEFAULT '',
'  `Text14` varchar(100) DEFAULT '',
'  `Text15` varchar(100) DEFAULT '',
'  `Text16` varchar(100) DEFAULT '',
'  `Text17` varchar(100) DEFAULT '',
'  `Text18` varchar(100) DEFAULT '',
'  `Text19` varchar(100) DEFAULT '',
'  `Text20` varchar(100) DEFAULT '',
'  `Text21` varchar(100) DEFAULT '',
'  `Text22` varchar(100) DEFAULT '',
'  `Text23` varchar(100) DEFAULT '',
'  `Text24` varchar(100) DEFAULT '',
'  `Text25` varchar(100) DEFAULT '',
'  `Text26` varchar(100) DEFAULT '',
'  `Text27` varchar(100) DEFAULT '',
'  `Text28` varchar(100) DEFAULT '',
'  `Text29` varchar(100) DEFAULT '',
'  `Text30` varchar(100) DEFAULT '',
'  `Text31` varchar(100) DEFAULT '',
'  `Text32` varchar(100) DEFAULT '',
'  `Text33` varchar(100) DEFAULT '',
'  `Text34` varchar(100) DEFAULT '',
'  `Text35` varchar(100) DEFAULT '',
'  `Text36` varchar(100) DEFAULT '',
'  `Text37` varchar(100) DEFAULT '',
'  `Text38` varchar(100) DEFAULT '',
'  `Text39` varchar(100) DEFAULT '',
'  `Text40` varchar(100) DEFAULT '',
'  `TextXL` varchar(1000) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;

Public Sub LeeArticulos()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset 'tabla articulo
    Dim Myrec2040 As dao.Recordset
    Dim MyRecTmp As dao.Recordset
    Dim MyRecAux As New ADODB.Recordset
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim nPrc As Double
    Dim sDesc As String
    Dim nLen As Long
    Dim ncont As Double
    Dim sParte As String
    Dim a As Variant
    Dim nContEvent As Integer

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from mbd.Items where id<>0 order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount

            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF

                nFam = MyRecSet.Fields("family")
                If nFam = 0 Then nFam = 1
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(nFam))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("codi_fam") = nFam
                    myRec.Fields("codi_ident") = 0 'MyRecSet.fields("counter")
                    myRec.Fields("secc_maqui") = 0 'MyRecSet.fields("counter")
                    myRec.Fields("posicion") = 0
                    myRec.Fields("index") = nFam
                    myRec.Fields("txt_fam") = "DPT. " & CStr(nFam)
                    myRec.Fields("imagen") = ""
                    myRec.Fields("etiqueta") = ""
                    myRec.Update
                End If
                myRec.Close
                Set myRec = Nothing
                'Set myRec = Mybase.OpenRecordset("select * from familias where codi_fam=" & CStr(nFam))
                'If myRec.EOF Then
                '    myRec.AddNew
                '    myRec.fields("codi_fam") = nFam
                '    myRec.fields("codi_ident") = MyRecSet.fields("counter")
                '    myRec.fields("secc_maqui") = MyRecSet.fields("counter")
                '    myRec.fields("index") = 0
                '    myRec.Update
                'End If
                'myRec.Close
                'Set myRec = Nothing
                MyRecAux.Open "Select * from mbd.ItemRates where id=" & CStr(MyRecSet.Fields("id")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    EstadoBal = 1000
                    canCelar = 1
                    Call frmControl.Balanza_ComError(0)
                Else
                    If MyRecAux.EOF Then
                        nPrc = 0
                    Else
                        nPrc = MyRecAux.Fields("Rate0")
                    End If
                End If
                MyRecAux.Close
                Set MyRecAux = Nothing
                
                Set myRec = Mybase.OpenRecordset("select * from articulo where codigo=" & CStr(MyRecSet.Fields("id")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                nSec = MyRecSet.Fields("counter")
                nPlu = MyRecSet.Fields("plu")
                myRec.Fields("codi_ident") = CodiIdent_FS(nSec)
                If nSec = 0 Then
                    myRec.Fields("secc_maqui") = CodiIdent_FS(nSec)
                Else
                    myRec.Fields("secc_maqui") = nSec
                End If
                myRec.Fields("codi_sub") = 0
                myRec.Fields("codi_fam") = nFam
                myRec.Fields("plu") = nPlu
                myRec.Fields("codigo") = MyRecSet.Fields("id")
                If MyRecSet.Fields("saleform") = 1 Then
                    myRec.Fields("codi_pes") = "W"
                Else
                    myRec.Fields("codi_pes") = "U"
                End If
                
                
                myRec.Fields("tipo_iva") = MyRecSet.Fields("vat")
                
                If UsaEuro Then
                    myRec.Fields("euros") = nPrc
                Else
                    myRec.Fields("precio") = nPrc
                End If
                myRec.Fields("prc3") = Format(Now, "yymmddhhmm")
                myRec.Fields("tara") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("tara_envasado") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("caducidad") = MyRecSet.Fields("Expiring")
                myRec.Fields("pref") = MyRecSet.Fields("BestBefore")
                'Atención! Touch Scale no dispone de:
                '-Peso Aproximado
                '-Merma
                '-Undades de Preselección
                myRec.Fields("poid") = MyRecSet.Fields("approxweight")
                myRec.Fields("fcb") = 1000 * MyRecSet.Fields("loss")
                myRec.Fields("ning") = MyRecSet.Fields("totpresel")
                'Atención! La etiqueta es por nombre... Se ha de crear un fichero
                ' en el PC con los nombres de etiquetas que existen en la Balanza
                ' para que se pueda asociar mediante un índice... O crear un campo
                ' en Articulo para contener el nombre...
                ' DE MOMENTO SE RECOGE NOMBRE DE ETIQUETA EN TRAN_EP (max 60 caracteres)
                myRec.Fields("label") = MyRecSet.Fields("label")
                '1.0.2 ShortName no se captura...
                myRec.Fields("des_plu1") = Mid(Trim(MyRecSet.Fields("name")), 1, 40) 'era shortname
                '''''''''''''''''''''''''''''''''
                'Atención! no se dispone de Imagen ni Color
                myRec.Fields("art_cb") = Trim(MyRecSet.Fields("ean13"))
                myRec.Fields("grupo_cons") = 0 'Grupo de conservación
                myRec.Fields("imagen") = Trim(MyRecSet.Fields("icon")) '& " _+_ " & _
                                         'Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("qlevel"))))
                myRec.Fields("ning") = MyRecSet.Fields("discount")
                
                If MyRecSet.Fields("mixed") <> 0 Then
                    myRec.Fields("lmix") = True
                Else
                    myRec.Fields("lmix") = False
                End If
                myRec.Fields("poid") = 1000 * (MyRecSet.Fields("mixedweight"))
                
                myRec.Fields("des_plu2") = Trim(Mid(MyRecSet.Fields("text01"), 1, 25))
                myRec.Fields("des_plu3") = Trim(Mid(MyRecSet.Fields("text02"), 36, 25))
                myRec.Fields("des_plu4") = Trim(Mid(MyRecSet.Fields("text03"), 61, 25))
                myRec.Fields("des_plu5") = Trim(Mid(MyRecSet.Fields("text04"), 86, 25))
                myRec.Fields("des_plu6") = Trim(Mid(MyRecSet.Fields("text05"), 111, 25))
                myRec.Fields("des_plu7") = Trim(Mid(MyRecSet.Fields("text06"), 136, 25))
                myRec.Fields("des_plu8") = Trim(Mid(MyRecSet.Fields("text07"), 161, 25))
                myRec.Fields("des_plu9") = Trim(Mid(MyRecSet.Fields("text08"), 186, 25))
                myRec.Fields("des_plu0") = Trim(Mid(MyRecSet.Fields("text09"), 211, 25))
                myRec.Fields("des_plux") = Trim(Mid(MyRecSet.Fields("text10"), 236, 25))
                myRec.Fields("des_plu11") = Trim(Mid(MyRecSet.Fields("text11"), 261, 25))
                myRec.Fields("des_plu12") = Trim(Mid(MyRecSet.Fields("text12"), 286, 25))
                myRec.Fields("des_plu13") = Trim(Mid(MyRecSet.Fields("text13"), 311, 25))
                myRec.Fields("des_plu14") = Trim(Mid(MyRecSet.Fields("text14"), 336, 25))
                myRec.Fields("des_plu15") = Trim(Mid(MyRecSet.Fields("text15"), 361, 25))
                myRec.Fields("des_plu16") = Trim(Mid(MyRecSet.Fields("text16"), 386, 25))
                myRec.Fields("des_plu17") = Trim(Mid(MyRecSet.Fields("text17"), 411, 25))
                myRec.Fields("des_plu18") = Trim(Mid(MyRecSet.Fields("text18"), 436, 25))
                myRec.Fields("des_plu19") = Trim(Mid(MyRecSet.Fields("text19"), 461, 25))
                myRec.Fields("des_plu20") = Trim(Mid(MyRecSet.Fields("text20"), 486, 25))

                If Not IsNull(Trim(MyRecSet.Fields("textxl"))) Then
                    Set Myrec2040 = Mybase.OpenRecordset("select * from text15 where codigo=" & CStr(MyRecSet.Fields("id")))
                    If Myrec2040.EOF Then
                        Myrec2040.AddNew
                    Else
                        Myrec2040.Edit
                    End If
                    Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("id"))
                    Myrec2040.Fields("plu") = nPlu
                    Myrec2040.Fields("mostrador") = nSec
                    Myrec2040.Fields("tran_txt2") = "-"
                    Myrec2040.Fields("text1") = MyRecSet.Fields("textxl")
                    Myrec2040.Update
                    Myrec2040.Close
                    Set Myrec2040 = Nothing
                End If
                
                Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("id")))
                If myRec.EOF Then
                    Myrec2040.AddNew
                Else
                    Myrec2040.Edit
                End If
                Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
                Myrec2040.Fields("plu") = nPlu
                Myrec2040.Fields("mostrador") = nSec
                Myrec2040.Fields("tran_txt2") = "-"
                Myrec2040.Fields("txt_21") = Trim(MyRecSet.Fields("text21"))
                Myrec2040.Fields("txt_22") = Trim(MyRecSet.Fields("text22"))
                Myrec2040.Fields("txt_23") = Trim(MyRecSet.Fields("text23"))
                Myrec2040.Fields("txt_24") = Trim(MyRecSet.Fields("text24"))
                Myrec2040.Fields("txt_25") = Trim(MyRecSet.Fields("text25"))
                Myrec2040.Fields("txt_26") = Trim(MyRecSet.Fields("text26"))
                Myrec2040.Fields("txt_27") = Trim(MyRecSet.Fields("text27"))
                Myrec2040.Fields("txt_28") = Trim(MyRecSet.Fields("text28"))
                Myrec2040.Fields("txt_29") = Trim(MyRecSet.Fields("text29"))
                Myrec2040.Fields("txt_30") = Trim(MyRecSet.Fields("text30"))
                Myrec2040.Fields("txt_31") = Trim(MyRecSet.Fields("text31"))
                Myrec2040.Fields("txt_32") = Trim(MyRecSet.Fields("text32"))
                Myrec2040.Fields("txt_33") = Trim(MyRecSet.Fields("text33"))
                Myrec2040.Fields("txt_34") = Trim(MyRecSet.Fields("text34"))
                Myrec2040.Fields("txt_35") = Trim(MyRecSet.Fields("text35"))
                Myrec2040.Fields("txt_36") = Trim(MyRecSet.Fields("text36"))
                Myrec2040.Fields("txt_37") = Trim(MyRecSet.Fields("text37"))
                Myrec2040.Fields("txt_38") = Trim(MyRecSet.Fields("text38"))
                Myrec2040.Fields("txt_39") = Trim(MyRecSet.Fields("text39"))
                Myrec2040.Fields("txt_40") = Trim(MyRecSet.Fields("text40"))
                Myrec2040.Update
                Myrec2040.Close
                Set Myrec2040 = Nothing
                    
                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                'Call frmControl.Balanza_ComOK
                nContEvent = nContEvent + 1
                If nContEvent >= 10 Then
                    nContEvent = 0
                    Do_Events
                End If
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
End Sub

'Families
'--------
'DROP TABLE IF EXISTS `mbd`.`Families`;
'CREATE TABLE  `mbd`.`Families` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  `Label` varchar(70) DEFAULT '0',
'  `Icon` varchar(100) DEFAULT '',
'  `IsBovine` int(11) DEFAULT '0',
'  `IsTrace` int(11) DEFAULT '0',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;

Public Sub LeeFamilias()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from mbd.Families order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast

            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(MyRecSet.Fields("id")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                'atención! V4 Touch Scale las Familias son Genéricas nos se asignan a sección.
                myRec.Fields("codi_fam") = MyRecSet.Fields("id")
                myRec.Fields("codi_ident") = 0 'MyRecSet.Fields("idseccion")
                myRec.Fields("secc_maqui") = 0 'MyRecSet.Fields("idseccion")
                If MyRecSet.Fields("isbovine") <> 0 Then
                    myRec.Fields("posicion") = 1
                    If MyRecSet.Fields("isbovine") = 0 Then
                        myRec.Fields("posicion") = 0
                    Else
                        myRec.Fields("posicion") = 1
                    End If
                Else
                    myRec.Fields("posicion") = 0 'MyRecSet.Fields("posicion")
                End If
                myRec.Fields("txt_fam") = Mid(Trim(MyRecSet.Fields("name")), 1, 35)
                myRec.Fields("etiqueta") = Trim(MyRecSet.Fields("label"))
                myRec.Fields("imagen") = Trim(MyRecSet.Fields("icon"))
                'atención! V4 Touch Scale en pcscale.families se almacena si es familia Vacuno y etiqueta
                '          a asignar a los artículos de la Familia. Se deberá recorrer la tabla de artículos
                '          para que los que coincidan con la Familia asignarles la Etiqueta. Por otra parte
                '          en Hydra se indica una sola familia de Vacuno en Configuración y se asigna número
                '          de etiqueta especial, habrá que asignar nombre en lugar de número.
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
        
        Else 'se crea la familia 1
        
            Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=1")
            If myRec.EOF Then
                myRec.AddNew
                myRec.Fields("codi_fam") = 1
                myRec.Fields("codi_ident") = 0
                myRec.Fields("secc_maqui") = 0
                myRec.Fields("posicion") = 0
                myRec.Fields("txt_fam") = "DPT. 1"
                myRec.Fields("imagen") = ""
                myRec.Fields("etiqueta") = ""
                myRec.Update
            End If
    
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
End Sub

'Vendors
'-------
'DROP TABLE IF EXISTS `mbd`.`Vendors`;
'CREATE TABLE  `mbd`.`Vendors` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  `Icon` varchar(100) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
Public Sub LeeVendedores()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from mbd.Vendors order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("id")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                
                myRec.Fields("ident_vend") = MyRecSet.Fields("id")
                myRec.Fields("codi_ident") = 0
                myRec.Fields("secc_maqui") = 0
                myRec.Fields("tec_vend") = 0
                myRec.Fields("tran_vend") = "*"
                myRec.Fields("nombre") = Trim(Mid(MyRecSet.Fields("name"), 1, 30))
                myRec.Fields("imagen") = Trim(MyRecSet.Fields("icon"))
                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
            MyRecSet.Close
            Set MyRecSet = Nothing
            myRec.Close
            Set myRec = Nothing
            
            
        End If
        
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
    
End Sub

'Counters
'--------
'DROP TABLE IF EXISTS `mbd`.`Counters`;
'CREATE TABLE  `mbd`.`Counters` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  `Text` varchar(100) DEFAULT '',
'  `Icon` varchar(100) DEFAULT '',
'  `Turn` int(11) DEFAULT '0',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;

Public Sub LeeSecciones()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim cm As String
    
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select * from mbd.Counters where id<>0 order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast


            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
        
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                If MyMinSec = 0 Then
                    MyMinSec = MyRecSet.Fields("id")
                End If
                Set myRec = Mybase.OpenRecordset("select * from seccion where codi_ident=" & CStr(MyRecSet.Fields("id")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                myRec.Fields("com") = 0
                myRec.Fields("sec_ip") = ""
                myRec.Fields("codi_ident") = MyRecSet.Fields("id")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("id")
                myRec.Fields("nombre") = Mid(MyRecSet.Fields("name"), 1, 25)
                myRec.Fields("descripcio") = Mid(MyRecSet.Fields("text"), 1, 25)
                'atención! V4 touch scale se dejará com GA y se verá si cambiar...
                myRec.Fields("enviardatos") = "GA"
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
End Sub

'Scales
'------
'DROP TABLE IF EXISTS `pcscale`.`scales`;
'CREATE TABLE  `pcscale`.`scales` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `MAC` varchar(17) NOT NULL,
'  `IP` varchar(16) NOT NULL,
'  `Number` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Type` char(1) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `MAC` (`MAC`)
') ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
'Tabla scales:
'.- Type: ‘S’, ‘N’, ‘A’
'    Balanza servidora (‘S’), esclava (‘N’), balanza no “conectada” (‘A’)
'Public Sub LeeEquipos()
'    Dim Mybase As DAO.Database
'    Dim myRec As DAO.Recordset
'    Dim MyRecSec As DAO.Recordset
'    Dim sData As String
'    Dim cm As String
'
'    On Error Resume Next
'    Set Mybase = DAO.OpenDatabase(Base_General)
'    MyRecSet.CursorLocation = adUseClient
'    MyRecSet.Open "select * from pcscale.scales order by number", MyConnObj, adOpenKeyset, adLockOptimistic
'    If Err.Number <> 0 Then
'        On Error GoTo 0
'        EstadoBal = 1000
'        canCelar = 1
'        Call frmControl.Balanza_ComError(0)
'    Else
'
'
'        If Not MyRecSet.EOF Then
'            frmMonitor.barra1.Value = 0
'            MyRecSet.MoveLast'
'
'
'            frmMonitor.barra1.Max = MyRecSet.Recordcount
'            MyRecSet.MoveFirst
'            'n = 1
'            Do While Not MyRecSet.EOF
'                Set myRec = Mybase.OpenRecordset("select * from equipos where numero_eqp=" & CStr(MyRecSet.fields("number")))
'                If myRec.EOF Then
'                    myRec.AddNew
'                Else
'                    myRec.Edit
'                End If
'                myRec.fields("numero_eqp") = MyRecSet.fields("number")
'                sData = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.fields("ip"))))
'                If GeneraTextoIP(sData) Then
'                    myRec.fields("prog_eqp") = sData
'                End If
'                'If MyRecSet.Fields("tipoconexion") = 1 Then
'                '    MyRec.Fields("modelo") = 0
'                'Else
'                    myRec.fields("modelo") = MyRecSet.fields("type")
'                'End If
'                myRec.fields("descripcio") = Mid(Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.fields("name")))), 1, 21)
'                If MyRecSet.fields("counter") <> 0 Then
'                    myRec.fields("codi_ident") = MyRecSet.fields("counter")
'                    myRec.fields("secc_maqui") = MyRecSet.fields("counter")
'                Else
'                    If MyMinSec <> 0 Then
'                        myRec.fields("codi_ident") = MyMinSec
'                        myRec.fields("secc_maqui") = MyMinSec
'                    Else
'                        myRec.fields("codi_ident") = 1
'                        myRec.fields("secc_maqui") = 1
'                    End If
'                End If
'                myRec.Update
'
'                'n = n + 1
'                MyRecSet.Movenext
'
'                Call MonitorBarra
'                Call frmControl.Balanza_ComOK
'
'                Do_Events
'
'                If canCelar = 1 Or CancelHumano = True Then GoTo salir
'
'            Loop
'
'
'        End If
'
'
'    End If
'salir:
'    On Error Resume Next
'    MyRecSet.Close
'    Set MyRecSet = Nothing
'    myRec.Close
'    Set myRec = Nothing
'    Mybase.Close
'    Set Mybase = Nothing
'
'FIN:
'    On Error GoTo 0
'End Sub

'VAT's
'DROP TABLE IF EXISTS `mbd`.`VATs`;
'CREATE TABLE  `mbd`.`VATs` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  `Percent` decimal(4,2) DEFAULT '0.00',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
Public Sub LeeIVAS()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset 'tabla articulo
    Dim MyRecAux As New ADODB.Recordset

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    'n = 0
    MyRecAux.CursorLocation = adUseClient
    MyRecAux.Open "Select * from mbd.VATs order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecAux.EOF Then
            MyRecAux.MoveLast
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecAux.Recordcount
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
             'MyRecSet.Open "Select * from dbo.tipoimpuesto where idtipoimpuesto=" & CStr(MyRecAux.Fields("idtipoimpuesto1")), MyConnObj, adOpenKeyset, adLockOptimistic
             'If Not MyRecSet.EOF Then
                Set myRec = Mybase.OpenRecordset("select * from iva where tipo=" & CStr(MyRecAux.Fields("id")))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("tipo") = MyRecAux.Fields("id")
                   'n = n + 1
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("id")
                   myRec.Fields("descripcion") = Trim(MyRecAux.Fields("name"))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                   'If n > 4 Then
                   '   GoTo Salir
                   'End If
                Else
                '    n = n + 1
                   myRec.Edit
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("id")
                   myRec.Fields("descripcion") = Trim(MyRecAux.Fields("name"))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                
                End If
             'End If
             MyRecAux.Movenext
             'MyRecSet.Close
             'Set MyRecSet = Nothing
             Call MonitorBarra
             Call frmControl.Balanza_ComOK

             Do_Events

             If canCelar = 1 Or CancelHumano = True Then GoTo salir

            Loop

        End If

    End If
    ''''''''''''''''''''''''''''''''

salir:
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    Mybase.Close
    Set Mybase = Nothing
fin:
    On Error GoTo 0

End Sub

'Paises
'DROP TABLE IF EXISTS `mbd`.`Countries`;
'CREATE TABLE  `mbd`.`Countries` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
Public Sub capturaPaises()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    On Error GoTo SalErrPaises
    MyRecSet.Open "Select * from mbd.Countries order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Mybase.Execute ("delete from paises")
        Set MyRecVar = Mybase.OpenRecordset("select * from paises")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            MyRecVar.AddNew
            MyRecVar.Fields("codigo") = MyRecSet.Fields("id")
            MyRecVar.Fields("nombre") = Mid(MyRecSet.Fields("name"), 1, 100)
            MyRecVar.Fields("valor") = Mid(MyRecSet.Fields("name"), 1, 100)
            MyRecVar.Fields("tran_pais") = " "
            MyRecVar.Fields("borrado") = False
            MyRecVar.Fields("modificado") = True
            MyRecVar.Fields("abrv_3") = "   "
            'On Error Resume Next
            MyRecVar.Fields("abrv_2") = "  "
            'On Error GoTo 0
            MyRecVar.Fields("tipo") = 1
            MyRecVar.Update
            MyRecSet.Movenext
        Loop
SalErrPaises:
        On Error GoTo 0
        MyRecVar.Close
        Set MyRecVar = Nothing
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'Fichas Vacuno
'DROP TABLE IF EXISTS `mbd`.`Bovine`;
'CREATE TABLE  `mbd`.`Bovine` (
'  `Id` int(11) DEFAULT NULL,
'  `EarTag` varchar(100) DEFAULT '',
'  `BirthLand` int(11) DEFAULT '0',
'  `sBirthLand` varchar(100) DEFAULT '',
'  `BirthDate` datetime DEFAULT '0000-00-00 00:00:00',
'  `BreedingLand` int(11) DEFAULT '0',
'  `sBreedingLand` varchar(100) DEFAULT '',
'  `SlaughterLand` int(11) DEFAULT '0',
'  `sSlaughterLand` varchar(100) DEFAULT '',
'  `SlaughterRSI` varchar(100) DEFAULT '',
'  `SlaughterDate` datetime DEFAULT '0000-00-00 00:00:00',
'  `ButcherLand` int(11) DEFAULT '0',
'  `sButcherLand` varchar(100) DEFAULT '',
'  `ButcherRSI` varchar(100) DEFAULT '',
'  `ButcherDate` datetime DEFAULT '0000-00-00 00:00:00',
'  `Lot` varchar(100) DEFAULT '',
'  `Date1` datetime DEFAULT '0000-00-00 00:00:00',
'  `Date2` datetime DEFAULT '0000-00-00 00:00:00',
'  `Date3` datetime DEFAULT '0000-00-00 00:00:00',
'  `Text1` varchar(100) DEFAULT '',
'  `Text2` varchar(100) DEFAULT '',
'  `Text3` varchar(100) DEFAULT '',
'  `Text4` varchar(100) DEFAULT '',
'  `Text5` varchar(100) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
Public Sub capturaFichas()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.Open "Select * from mbd.Bovine order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno where codigo=" & Chr(34) & CStr(MyRecSet.Fields("id")) & Chr(34))
            If MyRecVar.EOF Then
                MyRecVar.AddNew
            Else
                MyRecVar.Edit
            End If
                
            MyRecVar.Fields("codigo") = MyRecSet.Fields("id")
            MyRecVar.Fields("identificador") = Trim(MyRecSet.Fields("eartag"))
            MyRecVar.Fields("codnacimiento") = MyRecSet.Fields("birthland")
            MyRecVar.Fields("fecha_nac") = Format(MyRecSet.Fields("birthdate"), "dd/mm/yyyy")
            MyRecVar.Fields("codcrianza") = MyRecSet.Fields("breedingland")
            MyRecVar.Fields("codsacrificio") = MyRecSet.Fields("slaughterland")
            MyRecVar.Fields("rsisacrificio") = Trim(MyRecSet.Fields("SlaughterRSI"))
            MyRecVar.Fields("fecha_sac") = Format(MyRecSet.Fields("SlaughterDate"), "dd/mm/yyyy")
            MyRecVar.Fields("coddespiece") = MyRecSet.Fields("ButcherLand")
            MyRecVar.Fields("rsidespiece") = Trim(MyRecSet.Fields("ButcherRSI"))
            If IsNull(MyRecSet.Fields("lot")) Then
                MyRecVar.Fields("lote") = ""
            Else
                MyRecVar.Fields("lote") = Trim(MyRecSet.Fields("lot"))
            End If
            
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

Public Function consultaSeccMaquiBalanza(ByVal nBal) As Integer
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select secc_maqui from equipos where numero_eqp=" & CStr(nBal))
    If MyRecVar.EOF Then
        consultaSeccMaquiBalanza = 1
    Else
        consultaSeccMaquiBalanza = MyRecVar.Fields("secc_maqui")
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
End Function

'DROP TABLE IF EXISTS `mbd`.`BarcodeEan13`;
'CREATE TABLE  `mbd`.`BarcodeEan13` (
'  `Id` int(11) DEFAULT NULL,
'  `Type` varchar(1) DEFAULT '0',
'  `Number` int(11) DEFAULT '0',
'  `LocalSales` varchar(13) DEFAULT '',
'  `LocalSS` varchar(13) DEFAULT '',
'  `LocalPacket` varchar(13) DEFAULT '',
'  `LocalTotal` varchar(13) DEFAULT '',
'  `NetSales` varchar(13) DEFAULT '',
'  `NetSS` varchar(13) DEFAULT '',
'  `NetPacket` varchar(13) DEFAULT '',
'  `NetTotal` varchar(13) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
Public Sub capturaCB()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim nCodiIdent As Integer
Dim nSeccMaqui As Integer
Dim nC As Integer
Dim cm As String

    Set Mybase = dao.OpenDatabase(Base_General)
    ''''''''''''''''''''''''''
    cm = "delete from codbar"
    Mybase.Execute cm
    ''''''''''''''''''''''''''
    MyRecSet.Open "Select * from mbd.BarcodeEan13 where type<>'0'", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from codbar")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            If MyRecSet.Fields("type") = "2" Then
                nSeccMaqui = consultaSeccMaquiBalanza(MyRecSet.Fields("number"))
                nCodiIdent = nSeccMaqui
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where numero_eqp=" & CStr(MyRecSet.Fields("number")))
            Else
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where secc_maqui=" & CStr(MyRecSet.Fields("number")) & " and isnull(numero_eqp)")
                nSeccMaqui = MyRecSet.Fields("number")
                nCodiIdent = nSeccMaqui
            End If
            If MyRecVar.EOF Then
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.AddNew
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("type") = "2" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = MyRecSet.Fields("netsales")
                        MyRecVar.Fields("CB_SUP") = MyRecSet.Fields("Netsales")
                        MyRecVar.Fields("CB_MIX") = MyRecSet.Fields("NetTotal")
                        MyRecVar.Fields("CB_EV1") = MyRecSet.Fields("NetSS")
                        MyRecVar.Fields("CB_EV2") = MyRecSet.Fields("NetPacket")
                    Else
                        MyRecVar.Fields("CB_VNT") = MyRecSet.Fields("localsales")
                        MyRecVar.Fields("CB_SUP") = MyRecSet.Fields("localsales")
                        MyRecVar.Fields("CB_MIX") = MyRecSet.Fields("localTotal")
                        MyRecVar.Fields("CB_EV1") = MyRecSet.Fields("localSS")
                        MyRecVar.Fields("CB_EV2") = MyRecSet.Fields("localPacket")
                    End If
                    MyRecVar.Update
                Next nC
            Else
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.Edit
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("numbertype") = "S" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = MyRecSet.Fields("netsales")
                        MyRecVar.Fields("CB_SUP") = MyRecSet.Fields("Netsales")
                        MyRecVar.Fields("CB_MIX") = MyRecSet.Fields("NetTotal")
                        MyRecVar.Fields("CB_EV1") = MyRecSet.Fields("NetSS")
                        MyRecVar.Fields("CB_EV2") = MyRecSet.Fields("NetPacket")
                    Else
                        MyRecVar.Fields("CB_VNT") = MyRecSet.Fields("localsales")
                        MyRecVar.Fields("CB_SUP") = MyRecSet.Fields("localsales")
                        MyRecVar.Fields("CB_MIX") = MyRecSet.Fields("localTotal")
                        MyRecVar.Fields("CB_EV1") = MyRecSet.Fields("localSS")
                        MyRecVar.Fields("CB_EV2") = MyRecSet.Fields("localPacket")
                    End If
                    MyRecVar.Update
                Next nC
    
            End If
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'Clientes
'DROP TABLE IF EXISTS `mbd`.`customers`;
'CREATE TABLE  `mbd`.`customers` (
'  `Code` int(11) unsigned NOT NULL,
'  `CIF` varchar(16) NOT NULL DEFAULT '',
'  `Name` varchar(255) NOT NULL DEFAULT '',
'  `Address` varchar(255) DEFAULT '',
'  `City` varchar(255) DEFAULT '',
'  `ZIPCode` varchar(80) DEFAULT '',
'  `Country` varchar(80) DEFAULT '',
'  `Phone1` varchar(80) DEFAULT '',
'  `Phone2` varchar(80) DEFAULT '',
'  `Website` varchar(255) DEFAULT '',
'  `Email` varchar(80) DEFAULT '',
'  `Text1` varchar(255) DEFAULT '',
'  `Text2` varchar(255) DEFAULT '',
'  `Text3` varchar(255) DEFAULT '',
'  `Text4` varchar(255) DEFAULT '',
'  `Text5` varchar(255) DEFAULT '',
'  PRIMARY KEY (`Code`),
'  UNIQUE KEY `TIN` (`CIF`),
'  KEY `Name` (`Name`(19))
') ENGINE=MyISAM DEFAULT CHARSET=utf8;
Public Sub capturaClientes()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    On Error GoTo sal_CapturaClientes
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
    MyRecSet.Open "Select * from mbd.Customers  where code<>0 order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Set MyRecVar = Mybase.OpenRecordset("select * from cliente where ucase(trim(data3))='" & UCase(Trim(MyRecSet.Fields("cif"))) & "'")
            If MyRecVar.EOF Then
                MyRecVar.AddNew
            Else
                MyRecVar.Edit
            End If
            On Error Resume Next
            MyRecVar.Fields("cod_cli") = MyRecSet.Fields("code")
            MyRecVar.Fields("nom_cli") = Trim(MyRecSet.Fields("name"))
            MyRecVar.Fields("data1") = Trim(MyRecSet.Fields("address"))
            MyRecVar.Fields("data2") = Trim((MyRecSet.Fields("city")))
            MyRecVar.Fields("data3") = Trim(MyRecSet.Fields("cif"))
            MyRecVar.Fields("data4") = Trim(MyRecSet.Fields("phone1"))
            On Error GoTo 0
            On Error GoTo sal_CapturaClientes
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close
sal_CapturaClientes:
    If Err.Number <> 0 Then
        CadenadeLog "Error en CapturaClientes -->Err:" & CStr(Err.Number) & " desc.:" & Err.Description
    End If
    On Error GoTo 0
End Sub


Public Function verLF(ByVal sS As String) As String
Dim sR As String
    If lTransPlu Then
        sR = Replace(sS, Chr(94), Chr(10))
        sR = Replace(sR, " |", "<b>")
        verLF = Replace(sR, "|", "</b>")
    Else
        sR = Replace(sS, "/~", "</b>")
        sR = Replace(sR, "~", "<b>")
        sR = Replace(sR, Chr(94), Chr(10))
        verLF = Replace(sR, "|", Chr(10))
    End If
End Function
'//////////
'Envios SAM
'\\\\\\\\\\
Public Function Mue_Env_Gen(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean, Optional MiEnvio As Long) As Boolean
Dim Mybase As dao.Database
Dim myrecfam As dao.Recordset
Dim MyRecAux As New ADODB.Recordset
Dim cm As String

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINGEN
    End If
    On Error GoTo 0
    On Error GoTo FINGEN

    Set myrecfam = Mybase.OpenRecordset("select * from gen_sam where baja=true")
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
            Select Case myrecfam.Fields("tipo")
            Case "free"
                cm = "update mbd.Texts set free1='',free2='',free3='',free4='',free5='',free6='',free7='',free8='',free9='',free10='' where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                MyConnObj.Execute cm
            Case "cb14"
                cm = "delete from mbd.BarcodeEan14 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                MyConnObj.Execute cm
            Case "cb128"
                cm = "delete from mbd.BarcodeEan128 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                MyConnObj.Execute cm
            Case "cbCod"
                cm = "delete from mbd.BarcodeCode128 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                MyConnObj.Execute cm
            Case "cbgs1"
                cm = "delete from mbd.BarcodeGS1 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                MyConnObj.Execute cm
            End Select
            
            myrecfam.Delete
            
            If Not myrecfam.EOF Then myrecfam.Movenext
        Loop
        
    End If
    myrecfam.Close
    Set myrecfam = Nothing

    If Modificaciones Then
        Set myrecfam = Mybase.OpenRecordset("select * from gen_sam where tran_txt=' '")
    Else
        Set myrecfam = Mybase.OpenRecordset("select * from gen_sam")
    End If
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
            Select Case myrecfam.Fields("tipo")
            
            Case "free"
                MyRecAux.Open "select * from mbd.Texts where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINGEN
                End If
                On Error GoTo FINGEN
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                    MyRecAux.Fields("id") = CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                    MyRecAux.Fields("type") = Trim(myrecfam.Fields("destino"))
                    MyRecAux.Fields("number") = Trim(myrecfam.Fields("numero"))
                End If
                MyRecAux.Fields("free1") = Trim(myrecfam.Fields("txt01"))
                MyRecAux.Fields("free2") = Trim(myrecfam.Fields("txt02"))
                MyRecAux.Fields("free3") = Trim(myrecfam.Fields("txt03"))
                MyRecAux.Fields("free4") = Trim(myrecfam.Fields("txt04"))
                MyRecAux.Fields("free5") = Trim(myrecfam.Fields("txt05"))
                MyRecAux.Fields("free6") = Trim(myrecfam.Fields("txt06"))
                MyRecAux.Fields("free7") = Trim(myrecfam.Fields("txt07"))
                MyRecAux.Fields("free8") = Trim(myrecfam.Fields("txt08"))
                MyRecAux.Fields("free9") = Trim(myrecfam.Fields("txt09"))
                MyRecAux.Fields("free10") = Trim(myrecfam.Fields("txt10"))
                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            
            Case "cb14"
                MyRecAux.Open "select * from mbd.BarcodeEan14 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINGEN
                End If
                On Error GoTo FINGEN
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                    MyRecAux.Fields("id") = CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                    MyRecAux.Fields("type") = Trim(myrecfam.Fields("destino"))
                    MyRecAux.Fields("number") = Trim(myrecfam.Fields("numero"))
                End If
                MyRecAux.Fields("netsales") = Trim(myrecfam.Fields("txt01")) '!cb_vnt '1
                MyRecAux.Fields("nettotal") = Trim(myrecfam.Fields("txt02")) '!cb_mix '2
                MyRecAux.Fields("netss") = Trim(myrecfam.Fields("txt04")) '!cb_ev1 '4
                MyRecAux.Fields("netpacket") = Trim(myrecfam.Fields("txt05")) '!cb_ev2 '5
                MyRecAux.Fields("localsales") = Trim(myrecfam.Fields("txt06")) '!cb_vnt '6
                MyRecAux.Fields("localtotal") = Trim(myrecfam.Fields("txt07")) '!cb_mix '7
                MyRecAux.Fields("localss") = Trim(myrecfam.Fields("txt09")) '!cb_ev1 '9
                MyRecAux.Fields("localpacket") = Trim(myrecfam.Fields("txt10")) '!cb_ev2 '10
                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            Case "cb128"
                MyRecAux.Open "select * from mbd.BarcodeEan128 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINGEN
                End If
                On Error GoTo FINGEN
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                    MyRecAux.Fields("id") = CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                    MyRecAux.Fields("type") = Trim(myrecfam.Fields("destino"))
                    MyRecAux.Fields("number") = Trim(myrecfam.Fields("numero"))
                End If
                MyRecAux.Fields("IA00") = Trim(myrecfam.Fields("txt01"))
                MyRecAux.Fields("IA01") = Trim(myrecfam.Fields("txt02"))
                MyRecAux.Fields("IA02") = Trim(myrecfam.Fields("txt03"))
                MyRecAux.Fields("IA10") = Trim(myrecfam.Fields("txt04"))
                MyRecAux.Fields("IA11") = Trim(myrecfam.Fields("txt05"))
                MyRecAux.Fields("IA13") = Trim(myrecfam.Fields("txt06"))
                MyRecAux.Fields("IA15") = Trim(myrecfam.Fields("txt07"))
                MyRecAux.Fields("IA17") = Trim(myrecfam.Fields("txt08"))
                MyRecAux.Fields("IA21") = Trim(myrecfam.Fields("txt09"))
                MyRecAux.Fields("IA251") = Trim(myrecfam.Fields("txt10"))
                MyRecAux.Fields("IA30") = Trim(myrecfam.Fields("txt11"))
                MyRecAux.Fields("IA310") = Trim(myrecfam.Fields("txt12"))
                MyRecAux.Fields("IA330") = Trim(myrecfam.Fields("txt13"))
                MyRecAux.Fields("IA3922") = Trim(myrecfam.Fields("txt14"))
                MyRecAux.Fields("IA400") = Trim(myrecfam.Fields("txt15"))
                MyRecAux.Fields("IA414") = Trim(myrecfam.Fields("txt16"))
                MyRecAux.Fields("IA422") = Trim(myrecfam.Fields("txt17"))
                MyRecAux.Fields("IA423") = Trim(myrecfam.Fields("txt18"))
                MyRecAux.Fields("IA426") = Trim(myrecfam.Fields("txt19"))
                MyRecAux.Fields("IA7030") = Trim(myrecfam.Fields("txt20"))
                MyRecAux.Fields("IA7031") = Trim(myrecfam.Fields("txt21"))
                MyRecAux.Fields("IA8005") = Trim(myrecfam.Fields("txt22"))
                MyRecAux.Fields("IA950") = Trim(myrecfam.Fields("txt23"))

                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            Case "cbCod"
                MyRecAux.Open "select * from mbd.BarcodeCode128 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINGEN
                End If
                On Error GoTo FINGEN
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                    MyRecAux.Fields("id") = CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                    MyRecAux.Fields("type") = Trim(myrecfam.Fields("destino"))
                    MyRecAux.Fields("number") = Trim(myrecfam.Fields("numero"))
                End If
                MyRecAux.Fields("code128") = Trim(myrecfam.Fields("txt01"))
                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            
            Case "cbgs1"
                MyRecAux.Open "select * from mbd.BarcodeGS1 where id=" & CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINGEN
                End If
                On Error GoTo FINGEN
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                    MyRecAux.Fields("id") = CStr(myrecfam.Fields("destino") * 1000 + myrecfam.Fields("numero"))
                    MyRecAux.Fields("type") = Trim(myrecfam.Fields("destino"))
                    MyRecAux.Fields("number") = Trim(myrecfam.Fields("numero"))
                End If
                MyRecAux.Fields("IA00") = Trim(myrecfam.Fields("txt01"))
                MyRecAux.Fields("IA01") = Trim(myrecfam.Fields("txt02"))
                MyRecAux.Fields("IA02") = Trim(myrecfam.Fields("txt03"))
                MyRecAux.Fields("IA10") = Trim(myrecfam.Fields("txt04"))
                MyRecAux.Fields("IA11") = Trim(myrecfam.Fields("txt05"))
                MyRecAux.Fields("IA13") = Trim(myrecfam.Fields("txt06"))
                MyRecAux.Fields("IA15") = Trim(myrecfam.Fields("txt07"))
                MyRecAux.Fields("IA17") = Trim(myrecfam.Fields("txt08"))
                MyRecAux.Fields("IA21") = Trim(myrecfam.Fields("txt09"))
                MyRecAux.Fields("IA251") = Trim(myrecfam.Fields("txt10"))
                MyRecAux.Fields("IA30") = Trim(myrecfam.Fields("txt11"))
                MyRecAux.Fields("IA310") = Trim(myrecfam.Fields("txt12"))
                MyRecAux.Fields("IA330") = Trim(myrecfam.Fields("txt13"))
                MyRecAux.Fields("IA3922") = Trim(myrecfam.Fields("txt14"))
                MyRecAux.Fields("IA400") = Trim(myrecfam.Fields("txt15"))
                MyRecAux.Fields("IA414") = Trim(myrecfam.Fields("txt16"))
                MyRecAux.Fields("IA422") = Trim(myrecfam.Fields("txt17"))
                MyRecAux.Fields("IA423") = Trim(myrecfam.Fields("txt18"))
                MyRecAux.Fields("IA426") = Trim(myrecfam.Fields("txt19"))
                MyRecAux.Fields("IA7030") = Trim(myrecfam.Fields("txt20"))
                MyRecAux.Fields("IA7031") = Trim(myrecfam.Fields("txt21"))
                MyRecAux.Fields("IA8005") = Trim(myrecfam.Fields("txt22"))
                MyRecAux.Fields("IA950") = Trim(myrecfam.Fields("txt23"))

                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            
            End Select
            
            myrecfam.Movenext
        Loop
        
    End If
    myrecfam.Close
    Set myrecfam = Nothing



    Mue_Env_Gen = True
    Exit Function

FINGEN:
    CadenadeLog "Error en MUE_Env_Gen:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
    Mue_Env_Gen = False
    
End Function

Public Function Item_OrderInFamily_T(ByVal nFam As Long, ByVal nMethod As Long) As Long
'nFam == número de familia
'nMethod --> 0=Código Artículo, 1=Plu Artículo, 2=Name, 3=shortname
'SELECT code,family,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=3 order by shortname
Dim Resp As Long
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
Dim n As Long
    Resp = 9999
    On Error GoTo Sal_T
    Select Case nMethod
    Case 0
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM Items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 1
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM Items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by plu", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 2
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM Items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by name", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 3
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM Items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by shortname", MyConnObj, adOpenKeyset, adLockOptimistic
    Case Else
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM Items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    End Select
    Resp = 10000
    n = 0
    If Not MyRecDes.EOF Then
        MyRecDes.MoveFirst
        Do While Not MyRecDes.EOF
            Resp = 10001
            cm = "update Items set familyorder=" & CStr(MyRecDes.Fields("rank")) & " where code=" & CStr(MyRecDes.Fields("code")) & " and family=" & CStr(nFam)
            CadenadeLog cm
            MyConnObj.Execute cm
            n = n + 1
            MyRecDes.Movenext
        Loop
    End If
    Resp = 10002
    MyRecDes.Close
    Set MyRecDes = Nothing
Sal_T:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'Resp = 10000
    Else
        Resp = n
    End If
    Item_OrderInFamily_T = Resp
End Function

''''''''''''''''''''''
'ENVIO DE ARTÍCULOS...
''''''''''''''''''''''
'Se envían:
'   IVAS
'   Familias Altas/bajas
'   Artículos Altas/bajas
'   Teclas (Siempre) Altas/bajas
'Atención PLU MIXTO
Public Function MUE_EnviaArticuloNW(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean, Optional MiCodigo As Long) As Boolean
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
Dim myRecTmp1 As dao.Recordset
Dim myrecfam As dao.Recordset
Dim MyRecFam1 As dao.Recordset
Dim MyRecSec As dao.Recordset
Dim nD As Double
Dim MyRecAux As New ADODB.Recordset
Dim MyRecAux1 As New ADODB.Recordset
Dim MyRecFamSec As New ADODB.Recordset
Dim cm As String
Dim nf As Integer
Dim nL As Integer
Dim lEsAlta As Boolean
Dim nC(99) As Long
Dim ncont As Integer
Dim Resp As Long
Dim nPrice As Double
Dim nContPrecios As Integer
Dim nIdFinal As Double
Dim sColorTmp As String
Dim PosiRel As Long
Dim lEnvImage As Boolean
Dim nContEvnt As Integer
Dim s(5) As String
Dim lEsTest As Boolean
Dim lNewTEXT As Boolean
Dim cm1 As String
Dim sBCB As String

CadenadeLog "Entrada en Envio de Articulos..."
'''''''''''''''''''
'Call Corregir_Nulos
'''''''''''''''''

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINARTICULO
    End If
    On Error GoTo 0
    'On Error Resume Next
    'Set myRecTmp1 = Mybase.OpenRecordset("select * from text15")
    'If Err.Number <> 0 Then
    '    lNewTEXT = False
    'Else
    '    lNewTEXT = True
    'End If
    'On Error GoTo 0
    'myRecTmp1.Close
    'Set myRecTmp1 = Nothing
    
'PosiRel = Item_OrderInFamily_T(1, 0)
'PosiRel = Item_OrderInFamily_T(2, 0)
'PosiRel = Item_OrderInFamily_T(3, 0)
'PosiRel = Item_OrderInFamily_T(4, 0)
'PosiRel = Item_OrderInFamily_T(5, 0)
'PosiRel = Item_OrderInFamily_T(6, 0)
'PosiRel = Item_OrderInFamily_T(7, 0)
'PosiRel = Item_OrderInFamily_T(9, 0)


    On Error GoTo FINARTICULO
''''''''''''''''''''''''
'    Call capturaClientes
'''''''PROGRAMACIÓN IVA's

CadenadeLog "Verificación IVA..."
'DROP TABLE IF EXISTS `mbd`.`VATs`;
'CREATE TABLE  `mbd`.`VATs` (
'  `Id` int(11) DEFAULT NULL,
'  `Name` varchar(100) DEFAULT '',
'  `Percent` decimal(4,2) DEFAULT '0.00',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;
    Set myRecTmp1 = Mybase.OpenRecordset("select * from iva")
    If Not myRecTmp1.EOF Then
        myRecTmp1.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = myRecTmp1.Recordcount
        myRecTmp1.MoveFirst
        Do While Not myRecTmp1.EOF
            MyRecAux.Open "select * from mbd.VATs where code=" & CStr(myRecTmp1.Fields("tipo")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            End If
            
            MyRecAux.Fields("code") = myRecTmp1.Fields("tipo")
            MyRecAux.Fields("name") = Trim(myRecTmp1.Fields("descripcion"))
            MyRecAux.Fields("percent") = myRecTmp1.Fields("valor")
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            myRecTmp1.Movenext
        
            Call MonitorBarra
        
            'Do_Events '--> ver
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 20 Then
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
        
        Loop
    
    End If
    myRecTmp1.Close
    Set myRecTmp1 = Nothing
'''''''''''''''''''''''''''''''

'!!!!!!!!!!!!!!!!!!!!!!!!
'CadenadeLog "Baja Familias Fam_Code..."
'    Set myrecfam = Mybase.OpenRecordset("select * from fam_code where posicion=1525")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
'    If Not myrecfam.EOF Then
'
'        myrecfam.MoveFirst
'
'        Do While Not myrecfam.EOF
'
'            cm = "delete from mbd.Families where code=" & CStr(myrecfam.Fields("codi_fam"))
'            MyConnObj.Execute cm
'            cm = "delete from mbd.Items where family=" & CStr(myrecfam.Fields("codi_fam"))
'            MyConnObj.Execute cm
'            '1.5.7
'            cm = "delete from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
'            MyConnObj.Execute cm
'            cm = "delete from familias where codi_fam=" & CStr(myrecfam.Fields("codi_fam"))
'            Mybase.Execute cm
'            ''''''
'
'            myrecfam.Delete
'
'            If Not myrecfam.EOF Then myrecfam.Movenext
'        Loop
'
'    End If
'    myrecfam.Close
'    Set myrecfam = Nothing
'    ''''''''''''''''''''
    
CadenadeLog "Alta familias Fam_Code..."
'DROP TABLE IF EXISTS `mbd`.`families`;
'CREATE TABLE  `mbd`.`families` (
'  `Code` int(11) unsigned NOT NULL,
'  `Name` varchar(100) NOT NULL DEFAULT '',
'  `Icon` varchar(127) DEFAULT '',
'  `IsBovine` tinyint(1) DEFAULT '0',
'  PRIMARY KEY (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=MyISAM DEFAULT CHARSET=utf8;
    Set myrecfam = Mybase.OpenRecordset("select * from fam_code order by codi_fam")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
            
            lEsAlta = False
            
            On Error Resume Next
            MyRecAux.Open "select * from mbd.Families where code=" & CStr(myrecfam.Fields("codi_fam")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                lEsAlta = True
            End If
            
            If lEsAlta Then
                CadenadeLog "Se detecta familia:" & CStr(myrecfam.Fields("codi_fam")) & " No existente en Tabla Balanza..."
                MyRecAux.AddNew
            End If
            
            MyRecAux.Fields("code") = myrecfam.Fields("codi_fam")
            If IsNull(myrecfam.Fields("txt_fam")) Then
                MyRecAux.Fields("name") = "FAM. 1" & CStr(myrecfam.Fields("codi_fam"))
            Else
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("txt_fam")))
            End If
            On Error Resume Next
            If myrecfam.Fields("etiqueta") <> "" Then
                MyRecAux.Fields("label") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("etiqueta"))) 'era imagen
            Else
                If lEsAlta Then MyRecAux.Fields("label") = ""
            End If
            If myrecfam.Fields("imagen") <> "" Then
                MyRecAux.Fields("icon") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("imagen"))) 'era imagen
            Else
                If lEsAlta Then MyRecAux.Fields("icon") = ""
            End If
            
            If myrecfam.Fields("posicion") = 1 Then
                MyRecAux.Fields("isbovine") = 1
            Else
                MyRecAux.Fields("isbovine") = 0
            End If
            
            'CadenadeLog "Se graba..."
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            myrecfam.Movenext
            Do_Events
        Loop
    
        
    End If

    Do_Events
    myrecfam.Close
    Set myrecfam = Nothing
    
CadenadeLog "Alta familias Fam_Code Finalizada..."
'DROP TABLE IF EXISTS `mbd`.`CounterFamilies`;
'CREATE TABLE  `mbd`.`CounterFamilies` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `CounterCode` int(11) NOT NULL,
'  `FamilyIndex` int(11) NOT NULL DEFAULT '1',
'  `FamilyCode` int(11) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `idcofam` (`CounterCode`,`FamilyIndex`),
'  CONSTRAINT `fkcofam` FOREIGN KEY (`CounterCode`) REFERENCES `Counters` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
'    If lUpper Or Dir(App.Path & "\cambioposicion") <> "" Then
'        If (Not Modificaciones) Or (Dir(App.Path & "\cambioposicion") <> "") Then
'            cm = "delete from mbd.CounterFamilies"
'            MyConnObj.Execute cm
'        End If
'        If Dir(App.Path & "\cambioposicion") <> "" Then
'            Kill App.Path & "\cambioposicion"
'        End If
'
'        Set myrecfam = Mybase.OpenRecordset("select * from familias")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
'        If Not myrecfam.EOF Then
'            myrecfam.MoveFirst
'            Do While Not myrecfam.EOF
'                If myrecfam.Fields("borrado") = True Then
'                    '1.5.7
'                    cm = "delete from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
'                    MyConnObj.Execute cm
'                    ''''''
'                    myrecfam.Delete
'                Else
'                    'asignación de indice
'                    MyRecAux.Open "select * from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
'                    If MyRecAux.EOF Then
'                        MyRecAux.AddNew
'                        MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
'                        MyRecAux.Fields("familyindex") = 999999
'                        MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
'                        MyRecAux.Update
'                    End If
'                    MyRecAux.Close
'                    Set MyRecAux = Nothing
'
'                    Set MyRecFam1 = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(myrecfam.Fields("codi_fam")) & " and index<>0")
'                    If Not MyRecFam1.EOF Then
'                        'If lUpperNW Then
'                        '    cm = "update pcscale.counterfamilies set familyindex=" & CStr(MyRecFam1.Fields("index") Mod (MyRecFam.Fields("secc_maqui") * 100)) & " where familycode=" & CStr(MyRecFam.Fields("codi_fam"))
'                        'Else
'                            cm = "update mbd.CounterFamilies set familyindex=" & CStr(MyRecFam1.Fields("index")) & " where familycode=" & CStr(myrecfam.Fields("codi_fam"))
'                        'End If
'                        MyConnObj.Execute cm
'                    End If
'
'                    MyRecFam1.Close
'                    Set MyRecFam1 = Nothing
'
'                End If
'                If Not myrecfam.EOF Then myrecfam.Movenext
'            Loop
'        End If
'        myrecfam.Close
'        Set myrecfam = Nothing
'
'    Else
'
CadenadeLog "Alta familias 'Familias' a counter Families..."
        
        Set myrecfam = Mybase.OpenRecordset("select * from familias")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
        If Not myrecfam.EOF Then
            myrecfam.MoveFirst
            Do While Not myrecfam.EOF
                'If myrecfam.Fields("borrado") = True Or myrecfam.Fields("index") = 0 Then
                '    '1.5.7
                '    cm = "delete from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                '    MyConnObj.Execute cm
                '    ''''''
                '    myrecfam.Delete
                'Else
                    'asignación de indice
                    'cm = "delete from mbd.CounterFamilies where familyindex=" & CStr(myrecfam.Fields("index")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                    'MyConnObj.Execute cm
                    'cm = "delete from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                    'MyConnObj.Execute cm
                    
                    MyRecAux.Open "select * from mbd.CounterFamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                    End If
                    MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
                    MyRecAux.Fields("familyindex") = myrecfam.Fields("index")
                    MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
                    MyRecAux.Fields("department") = 1 'se asigna al dpto. 1 de momento
                    MyRecAux.Update

                    MyRecAux.Close
                    Set MyRecAux = Nothing
        
                'End If
                If Not myrecfam.EOF Then myrecfam.Movenext
            Loop
        End If
        myrecfam.Close
        Set myrecfam = Nothing
    
    'End If
    
CadenadeLog "Fin familias 'familias' a counterfamilies..."
'!!!!!!!!!!!!!!!!!!!!!!!!

    lEsAlta = False
    '''''''''''''''''''''''''
    Set MyRecVar = Mybase.OpenRecordset("select * from articulo where borrado=true order by codigo")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo CargaCadena(1409) & " " & CStr(MyRecVar.Fields("codigo"))
            cm = "delete from mbd.Items where code=" & CStr(MyRecVar.Fields("codigo"))
            MyConnObj.Execute cm
            cm = "delete from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error Resume Next
            cm = "delete from text15 where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error GoTo 0
            On Error Resume Next
            cm = "delete from teclas where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error GoTo 0
            Do_Events
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo where tran_plu<>'*' and codigo<>0 order by codigo")
        lEnvImage = False
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo where codigo<>0 order by codigo")
        lEnvImage = True
    End If
    
    If Not MyRecVar.EOF Then
        MyRecVar.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
            
        nContEvnt = 0
        frmMonitor.MostrarInfo CargaCadena(230) & " " & CStr(frmMonitor.barra1.Max)
        CadenadeLog "Artículos a enviar...:" & CStr(frmMonitor.barra1.Max)
        Do While Not MyRecVar.EOF
            
            'CadenadeLog "code:" & MyRecVar.Fields("codigo")
            lEsAlta = False
            On Error Resume Next

            MyRecAux.Open "select * from mbd.Items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                'MyRecAux.AddNew
                lEsAlta = True
            End If
            '
            'MyRecAux.Close
            'Set MyRecAux = Nothing
            '
            'USAR PARA TEXT
            'lNewTEXT = False
            cm1 = ""
            Set myRecTmp1 = Mybase.OpenRecordset("select * from text15 where codigo=" & CStr(MyRecVar.Fields("codigo")))
            If Not myRecTmp1.EOF Then
                cm1 = Replace(Trim(myRecTmp1.Fields("text1")), "'", ".")
            End If

            If lEsAlta Then
                cm = "insert into mbd.items (code,name,text) values (" & CStr(MyRecVar.Fields("codigo")) & ",'" & _
                    Replace(Trim(MyRecVar.Fields("des_plu1")), "'", ".") & "','" & cm1 & "')"
                    
                'cm = "replace into mbd.items (code,name,text1,text2,text3,text4,text5,text6,text7,text8,text9,text10) values (" & CStr(MyRecVar.Fields("codigo")) & ",'" & _
                '    Trim(MyRecVar.Fields("des_plu1")) & "','" & Trim(MyRecVar.Fields("des_plu2")) & "','" & Trim(MyRecVar.Fields("des_plu3")) & "','" & Trim(MyRecVar.Fields("des_plu4")) & _
                '    "','" & Trim(MyRecVar.Fields("des_plu5")) & "','" & Trim(MyRecVar.Fields("des_plu6")) & "','" & Trim(MyRecVar.Fields("des_plu7")) & "','" & Trim(MyRecVar.Fields("des_plu8")) & _
                '    "','" & Trim(MyRecVar.Fields("des_plu9")) & "','" & Trim(MyRecVar.Fields("des_plu0")) & "','" & Trim(MyRecVar.Fields("des_plux")) & "')"
                    
            Else
                cm = "update mbd.items set name='" & Replace(Trim(MyRecVar.Fields("des_plu1")), "'", ".") & "',text='" & cm1 & "' " & _
                     "where code=" & CStr(MyRecVar.Fields("codigo"))
            End If
            MyConnObj.Execute cm
            
            myRecTmp1.Close
            Set myRecTmp1 = Nothing
            
            If lEsAlta Then 'quitar cualdo se haga insert completo
                MyRecAux.Close
                Set MyRecAux = Nothing
                MyRecAux.Open "select * from mbd.Items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
            End If
            
'Rehacer RPI
'            MyRecAux.Fields("text11") = Trim(MyRecVar.Fields("des_plu11"))
'            MyRecAux.Fields("text12") = Trim(MyRecVar.Fields("des_plu12"))
'            MyRecAux.Fields("text13") = Trim(MyRecVar.Fields("des_plu13"))
'            MyRecAux.Fields("text14") = Trim(MyRecVar.Fields("des_plu14"))
'            MyRecAux.Fields("text15") = Trim(MyRecVar.Fields("des_plu15"))
'            MyRecAux.Fields("text16") = Trim(MyRecVar.Fields("des_plu16"))
'            MyRecAux.Fields("text17") = Trim(MyRecVar.Fields("des_plu17"))
'            MyRecAux.Fields("text18") = Trim(MyRecVar.Fields("des_plu18"))
'            MyRecAux.Fields("text19") = Trim(MyRecVar.Fields("des_plu19"))
'            MyRecAux.Fields("text20") = Trim(MyRecVar.Fields("des_plu20"))
'            Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
'            If Not myRecTmp1.EOF Then
'                MyRecAux.Fields("text21") = Trim(myRecTmp1.Fields("txt_21"))
'                MyRecAux.Fields("text22") = Trim(myRecTmp1.Fields("txt_22"))
'                MyRecAux.Fields("text23") = Trim(myRecTmp1.Fields("txt_23"))
'                MyRecAux.Fields("text24") = Trim(myRecTmp1.Fields("txt_24"))
'                MyRecAux.Fields("text25") = Trim(myRecTmp1.Fields("txt_25"))
'                MyRecAux.Fields("text26") = Trim(myRecTmp1.Fields("txt_26"))
'                MyRecAux.Fields("text27") = Trim(myRecTmp1.Fields("txt_27"))
'                MyRecAux.Fields("text28") = Trim(myRecTmp1.Fields("txt_28"))
'                MyRecAux.Fields("text29") = Trim(myRecTmp1.Fields("txt_29"))
'                MyRecAux.Fields("text30") = Trim(myRecTmp1.Fields("txt_30"))
'                MyRecAux.Fields("text31") = Trim(myRecTmp1.Fields("txt_31"))
'                MyRecAux.Fields("text32") = Trim(myRecTmp1.Fields("txt_32"))
'                MyRecAux.Fields("text33") = Trim(myRecTmp1.Fields("txt_33"))
'                MyRecAux.Fields("text34") = Trim(myRecTmp1.Fields("txt_34"))
'                MyRecAux.Fields("text35") = Trim(myRecTmp1.Fields("txt_35"))
'                MyRecAux.Fields("text36") = Trim(myRecTmp1.Fields("txt_36"))
'                MyRecAux.Fields("text37") = Trim(myRecTmp1.Fields("txt_37"))
'                MyRecAux.Fields("text38") = Trim(myRecTmp1.Fields("txt_38"))
'                MyRecAux.Fields("text39") = Trim(myRecTmp1.Fields("txt_39"))
'                MyRecAux.Fields("text40") = Trim(myRecTmp1.Fields("txt_40"))
'            End If
'            myRecTmp1.Close
'            Set myRecTmp1 = Nothing
            
'**************************
'DROP TABLE IF EXISTS `mbd`.`items`;
'CREATE TABLE  `mbd`.`items` (
'  `Code` int(11) unsigned NOT NULL,
'  `Name` varchar(127) NOT NULL DEFAULT '',
'  `Text` text,
'  `Family` int(11) NOT NULL DEFAULT '0',
'  `FamilyOrder` int(11) DEFAULT '0',
'  `OnKey` tinyint(1) DEFAULT '1',
'  `Price` decimal(7,2) NOT NULL DEFAULT '0.00',
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Price100g` tinyint(4) DEFAULT '0',
'  `PLU` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `VAT` decimal(1,0) unsigned DEFAULT '0',
'  `bcActive` int(2) NOT NULL DEFAULT '0',
'  `EAN8` varchar(8) DEFAULT NULL,
'  `EAN13` varchar(13) DEFAULT NULL,
'  `EAN14` varchar(14) DEFAULT NULL,
'  `CODE39` varchar(29) DEFAULT NULL,
'  `CODE128` varchar(29) DEFAULT NULL,
'  `Expiring` smallint(6) DEFAULT '0',
'  `Tare` decimal(7,3) DEFAULT '0.000',
'  `Icon` varchar(127) DEFAULT '',
'  PRIMARY KEY (`Code`),
'  KEY `Name` (`Name`(19)),
'  KEY `ide_fmor` (`Family`,`FamilyOrder`)
') ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
            'CadenadeLog "Código:" & CStr(MyRecVar.Fields("codigo")) & " Precio:" & CStr(MyRecVar.Fields("precio")) & " Euros:" & CStr(MyRecVar.Fields("euros"))
            MyRecAux.Fields("family") = MyRecVar.Fields("codi_fam")
'GoTo myNext
            If MyRecVar.Fields("codi_pes") = "W" Then
                MyRecAux.Fields("saleform") = 0
            Else
                MyRecAux.Fields("saleform") = 1
            End If
            MyRecAux.Fields("plu") = MyRecVar.Fields("plu")
            MyRecAux.Fields("counter") = MyRecVar.Fields("secc_maqui")
            '3.8.1
            'MyRecAux.Fields("section") = MyRecVar.Fields("secc_maqui")
            ''''''
            MyRecAux.Fields("vat") = MyRecVar.Fields("tipo_iva")
            'MyRecAux.Fields("discount") = MyRecVar.Fields("ning")

            If Trim(MyRecVar.Fields("art_cb")) <> "" Then
                sBCB = Trim(MyRecVar.Fields("art_cb"))
                If Len(sBCB) >= 13 Then
                    sBCB = Mid(sBCB, 1, 13)
                Else
                    sBCB = Mid("0000000000000", 1, 12 - Len(sBCB)) & sBCB
                    sBCB = sBCB & Format(DigitoControl(sBCB), "0")
                End If
                MyRecAux.Fields("ean13") = sBCB
                'sBCB = "1" & Mid(sBCB, 1, 12) & Format(DigitoControl(sBCB), "0")
                'MyRecAux.Fields("ean14") = sBCB
                MyRecAux.Fields("bcactive") = 2
            Else
                MyRecAux.Fields("ean13") = ""
                'MyRecAux.Fields("ean14") = ""
                MyRecAux.Fields("bcactive") = 0
            End If
            'Ha de ser code 128...
            'If Trim(MyRecVar.Fields("art_cb14")) <> "" Then
            '    sBCB = Trim(MyRecVar.Fields("art_cb14"))
            '    If Len(sBCB) >= 14 Then
            '        sBCB = Mid(sBCB, 1, 14)
            '    Else
            '        sBCB = Mid("00000000000000", 1, 13 - Len(sBCB)) & sBCB
            '        sBCB = sBCB & Format(DigitoControl(sBCB), "0")
            '    End If
            '    MyRecAux.Fields("ean14") = sBCB
            '    'sBCB = "1" & Mid(sBCB, 1, 12) & Format(DigitoControl(sBCB), "0")
            '    'MyRecAux.Fields("ean14") = sBCB
            '    MyRecAux.Fields("bcactive") = 3
            'Else
            '    MyRecAux.Fields("ean14") = ""
            '    'MyRecAux.Fields("ean14") = ""
            'End If
            'FALTA EAN 8
            
            MyRecAux.Fields("expiring") = MyRecVar.Fields("caducidad")
            'MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
            MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
            'If Trim(MyRecVar.Fields("label2")) <> "" Then
            '    MyRecAux.Fields("label2") = Trim(MyRecVar.Fields("label2"))
            'End If
            If Trim(MyRecVar.Fields("tran_ep")) <> "" Then
                MyRecAux.Fields("xmllabel") = Trim(MyRecVar.Fields("tran_ep"))
            Else

                Select Case MyRecVar.Fields("etq")
                Case 0
                    MyRecAux.Fields("xmllabel") = ""
                Case 255
                    MyRecAux.Fields("xmllabel") = ""
                Case Else
                    MyRecAux.Fields("xmllabel") = CStr(MyRecVar.Fields("etq")) & ".xml"
                End Select

            End If
            '" _+_ "
            If Not IsNull(MyRecVar.Fields("imagen")) Then
                s(1) = separaFich(Trim(MyRecVar.Fields("imagen")))
                If Not IsNull(s(1)) Then
                    If Trim(s(1)) <> "" Then
                        MyRecAux.Fields("icon") = Trim(s(1))
                    Else
                        If lEsAlta Then MyRecAux.Fields("icon") = ""
                    End If
                Else
                    If lEsAlta Then MyRecAux.Fields("icon") = ""
                End If
            Else
                If lEsAlta Then MyRecAux.Fields("icon") = ""
            End If
            'If MyRecVar.Fields("lmix") Then
            '    MyRecAux.Fields("mixed") = 1
            'Else
            '    MyRecAux.Fields("mixed") = 0
            'End If
            'MyRecAux.Fields("mixedweight") = (MyRecVar.Fields("poid") / 1000)
            ''''
            If Dir(App.Path & "\trataonkey.ver") <> "" Then
                If MyRecVar.Fields("onkey") Then
                    MyRecAux.Fields("onkey") = 1
                Else
                    MyRecAux.Fields("onkey") = 0
                End If
            Else
                MyRecAux.Fields("onkey") = 1
            End If
            ''''
            MyRecAux.Fields("price") = CDbl(MyRecVar.Fields("euros")) 'Replace("'" & MyRecVar.Fields("euros") & "'", ",", ".")
            
myNext:
            MyRecAux.Update

            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'CadenadeLog "code:" & MyRecVar.Fields("codigo")

            MyRecVar.Edit
            MyRecVar.Fields("tran_plu") = "*"
            MyRecVar.Update


''DROP TABLE IF EXISTS `mbd`.`ItemPricerates`;
''CREATE TABLE  `mbd`.`ItemPricerates` (
''  `Id` int(11) NOT NULL AUTO_INCREMENT,
''  `Code` int(11) NOT NULL,
''  `Rate` decimal(1,0) NOT NULL DEFAULT '1',
''  `Price` decimal(8,2) NOT NULL DEFAULT '0.00',
''  `Step` decimal(5,3) NOT NULL DEFAULT '0.000',
''  PRIMARY KEY (`Id`),
''  UNIQUE KEY `idx_rate` (`Code`,`Rate`,`Step`),
''  KEY `Rate` (`Rate`),
''  CONSTRAINT `ItemPricerates_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `Items` (`Code`) ON DELETE CASCADE,
''  CONSTRAINT `ItemPricerates_ibfk_2` FOREIGN KEY (`Rate`) REFERENCES `Pricerates` (`Code`) ON DELETE CASCADE
'') ENGINE=InnoDB DEFAULT CHARSET=utf8;
''DROP TABLE IF EXISTS `mbd`.`Pricerates`;
''CREATE TABLE  `mbd`.`Pricerates` (
''  `Code` decimal(1,0) NOT NULL,
''  `Name` varchar(64) NOT NULL DEFAULT '',
''  PRIMARY KEY (`Code`)
'') ENGINE=InnoDB DEFAULT CHARSET=utf8;
            
'            For nf = 0 To 8
'                MyRecAux1.Open "Select * from mbd.Pricerates where code=" & CStr(nf + 1), MyConnObj, adOpenKeyset, adLockOptimistic
'                If MyRecAux1.EOF Then
'                    MyRecAux1.AddNew
'                End If
'                MyRecAux1.Fields("name") = "Tarif. " & CStr(nf + 1)
'                MyRecAux1.Fields("code") = nf + 1
'                MyRecAux1.Update
'                MyRecAux1.Close
'                Set MyRecAux1 = Nothing
'
'                MyRecAux1.Open "Select * from mbd.ItemPricerates where code=" & CStr(MyRecVar.Fields("codigo") & " and rate=" & CStr(nf + 1)), MyConnObj, adOpenKeyset, adLockOptimistic
'                If MyRecAux1.EOF Then
'                    MyRecAux1.AddNew
'                End If
'                MyRecAux1.Fields("code") = MyRecVar.Fields("codigo")
'                MyRecAux1.Fields("rate") = nf + 1
'                Select Case nf
'                'Case 0
'                '    If UsaEuro Then
'                '        MyRecAux1.Fields("price") = MyRecVar.Fields("euros")
'                '    Else
'                '        MyRecAux1.Fields("price") = MyRecVar.Fields("precio")
'                '    End If
'                Case 0
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar1")
'                Case 1
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar2")
'                Case 2
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar3")
'                Case 3
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar4")
'                Case 4
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar5")
'                Case 5
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar6")
'                Case 6
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar7")
'                Case 7
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar8")
'                Case 8
'                    MyRecAux1.Fields("price") = MyRecVar.Fields("tar9")
'                End Select
'                MyRecAux1.Update
'                MyRecAux1.Close
'                Set MyRecAux1 = Nothing
'            Next nf
'
            MyRecVar.Movenext
'
'            'Call MonitorBarra'
'
'            'Call frmControl.Balanza_ComOK
'
'            'Do_Events '--> ver
'**************************************
            
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 100 Then
                Call MonitorBarra
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirArt
            
        Loop
        
        MyRecVar.Close
        Set MyRecVar = Nothing
        
    End If
    
    
        
    nContEvnt = 0
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and ((tran_tec<>'*') or isnull(tran_tec))")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo "Programación Especial Tecla Código:" & " " & CStr(MyRecVar.Fields("codigo"))
            
            MyRecAux.Open "select * from mbd.Items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
            
            If Not MyRecAux.EOF Then
                If MyRecVar.Fields("tabla") <> 0 Then
                    MyRecAux.Fields("family") = MyRecVar.Fields("tabla")
                    MyRecAux.Fields("onkey") = 1
                    MyRecAux.Fields("familyorder") = MyRecVar.Fields("codi_tec")
                    MyRecAux.Update
                End If
                'Do_Events
                nContEvnt = nContEvnt + 1
                If nContEvnt >= 20 Then
                    nContEvnt = 0
                    DoEvents
                    Sleep (1)
                End If
                
            End If
            MyRecAux.Update
            
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_tec") = "*"
            MyRecVar.Update
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    
'''''''''''''''''''''''''''''''

SalirArt:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing

    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    
CadenadeLog "Fin envio Artículos."
    
    MUE_EnviaArticuloNW = True
    Exit Function
FINARTICULO:
    CadenadeLog "Error en MUE_EnviaArticuloNW:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
    MUE_EnviaArticuloNW = False
End Function

'
Public Sub MUE_EnviarFichasVacuno(ByVal NSistema As Integer, ByVal Modificaciones As Boolean, ByVal Marcar As Boolean, Optional MyCod As Long)
'Dim Crianza As Long
'Dim Nacimiento As Long
'Dim Produccion As Long
'Dim sCrianza As String
'Dim sNacimiento As String
'Dim sProduccion As Long
'Dim Sacrificio As Long
'Dim sSacrificio As String
'Dim Despiece As Long
'Dim sDespiece As Long
'Dim sSexo As String
'Dim sRaza As String
'Dim sCategoria As String
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nC1 As Integer
Dim cm As String
'Dim nLast As Double

    If MyCod = 0 Then frmMonitor.MostrarInfo CargaCadena(619) '"Envío de Fichas Vacuno"
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
        Set MyRecVar = Mybase.OpenRecordset("select * from paises where codigo<>0")
        If Not MyRecVar.EOF Then
            MyRecVar.MoveFirst
            Do While Not MyRecVar.EOF
                MyRecDes.Open "Select * from mbd.Countries where id=" & CStr(MyRecVar.Fields("codigo")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    EstadoBal = 1000
                    canCelar = 1
                    Call frmControl.Balanza_ComError(0)
                    GoTo salir
                Else
                    If MyRecDes.EOF Then
                        MyRecDes.AddNew
                        MyRecDes.Fields("id") = MyRecVar.Fields("codigo")
                        MyRecDes.Fields("name") = Trim(MyRecVar.Fields("nombre"))
                    Else
                        MyRecDes.Fields("name") = Trim(MyRecVar.Fields("nombre"))
                    End If
                    
                    MyRecDes.Update
                    MyRecDes.Close
                    Set MyRecDes = Nothing
                    
                    frmMonitor.MostrarDato "Country:" & " " & Format(MyRecVar.Fields("codigo"), "000") & " --> OK"
                End If
                MyRecVar.Movenext
            Loop
        End If
        MyRecVar.Close
        Set MyRecVar = Nothing
        
        Do_Events
        If MyCod = 0 Then
            If Modificaciones Then
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno where  (tranficha<>'*' OR tranficha is null)")
            Else
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno ")
            End If
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from fichavacuno where borrado=false and codigo=" & Chr(34) & MyCod & Chr(34))
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    If !borrado = False Then
                        'Crianza = !codcrianza
                        'On Error Resume Next
                        'sCrianza = Trim(!valorpcrianza)
                        'Nacimiento = !codnacimiento
                        'sNacimiento = Trim(!valorpnacimiento)
                        'Produccion = !codproduccion
                        'sProduccion = Trim(!valorpproduccion)
                        'Sacrificio = !codpsacrificio
                        'sSacrificio = Trim(!valorpsacrificio)
                        'Despiece = !codpdespiece
                        'sDespiece = Trim(!valorpdespiece)
                        'On Error GoTo 0
                        'If !codcrianza = 0 And !codnacimiento = 0 Then
                        '    Crianza = Produccion
                        '    sCrianza = sProduccion
                        '    Nacimiento = Produccion
                        '    sNacimiento = sProduccion
                        'End If
                        'Valores paises, categoria, raza, sexo
                        
                        ''''''''''''''''''''''''''''''''''''''
'Id (int 11) EarTag (varchar 100)
'BirthLand (int 11) sBirthLand (varchar 100) BirthDate (datetime)
'BreedingLand (int 11) sBreedingLand (varchar 100)
'SlaughterLand (int 11) sSlaughterLand (varchar 100)
'SlaughterRSI (varchar 100) SlaughterDate (datetime)
'ButcherLand (int 11) sButcherLand (varchar 100) ButcherRSI (varchar 100) ButcherDate (datetime)
'Lot (varchar 100)
'Date1 (datetime) Date2 (datetime) Date3 (datetime)
'Text1 (varchar 100) Text2 (varchar 100) Text3 (varchar 100) Text4 (varchar 100) Text5 (varchar 100)
                        ''''''''''''''''''''''''''''''''''''''
                        MyRecDes.Open "Select * from mbd.Bovine where id=" & CStr(!codigo), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            If MyRecDes.EOF Then

                                MyRecDes.AddNew
                                MyRecDes.Fields("id") = !codigo
                                MyRecDes.Fields("eartag") = Trim(!Identificador)
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterRSI") = Trim(!rsisacrificio)
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherRSI") = Trim(!rsidespiece)
                                
                                MyRecDes.Fields("lot") = Trim(!lote)
                                    
                            
                            Else
                                   
                                MyRecDes.Fields("id") = !codigo
                                MyRecDes.Fields("eartag") = Trim(!Identificador)
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterRSI") = Trim(!rsisacrificio)
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherRSI") = Trim(!rsidespiece)
                                MyRecDes.Fields("lot") = Trim(!lote)
                                   
                               
                            End If
                            
                            MyRecDes.Update
                            MyRecDes.Close
                            Set MyRecDes = Nothing
                            
                            frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo, "000") & " --> OK"
                            'Edit_Record Registro
                            .Edit
                            .Fields("tranficha") = "*"
                            .Update
    
                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        cm = "delete from pcscale.Bovine where id=" & CStr(!codigo)
                        MyConnObj.Execute cm
                        ''''''''''''''''''''
                        frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tranficha") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirFicha:
                    
                Loop
            End If
        End With
    
SalirFicha:
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing

salir:

End Sub

Public Sub MUE_EnviaCLientes(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean)
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim cm As String

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCLIENTES
    End If
    On Error GoTo 0
    On Error GoTo FINCLIENTES
    Set MyRecVar = Mybase.OpenRecordset("select * from cliente where borrado=true")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo CargaCadena(1410) & CStr(MyRecVar.Fields("cod_cli"))
            cm = "delete from mbd.Customers where id=" & CStr(MyRecVar.Fields("cod_cli"))
            MyConnObj.Execute cm
            cm = "delete from mbd.Customers where trim(upper(tin))='" & Trim(UCase(MyRecVar.Fields("data3")) & "'")
            MyConnObj.Execute cm
            
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente where tran_cli<>'*'")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            
            MyRecDes.Open "Select * from mbd.Customers where trim(upper(tin))='" & Trim(UCase(MyRecVar.Fields("data3")) & "'"), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINCLIENTES
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            
            frmMonitor.MostrarInfo CargaCadena(1411) & " " & CStr(MyRecVar.Fields("cod_cli"))
            
            MyRecDes.Fields("id") = MyRecVar.Fields("cod_cli")
            MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nom_cli")))
            MyRecDes.Fields("address") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("data1")))
            MyRecDes.Fields("city") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data2")))
            'MyRecDes.Fields("web") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data3")))
            MyRecDes.Fields("web") = ""
            MyRecDes.Fields("tin") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data3")))
            MyRecDes.Fields("phone") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data4")))
            'MyRecDes.fields("text1") = ""
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_cli") = "*"
            MyRecVar.Update
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCliente
            
        Loop
    
    
    End If
SalirCliente:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCLIENTES:
    CadenadeLog "Error en MUE_EnviaClientes:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
End Sub

Public Function Ii(ByVal sS As String) As String
Dim nPos As Integer
Dim sR As String
    sR = sS
    nPos = InStr(1, sR, "I")
    If nPos <> 0 Then
        Do While InStr(1, sR, "I") <> 0
            nPos = InStr(1, sR, "I")
            sR = Mid(sR, 1, nPos - 1) & "i" & Mid(sR, nPos + 1)
        Loop
    End If
    Ii = sR
End Function

Public Sub MUE_EnviaCodigosBarras(NSistema As Integer, Modificaciones As Boolean, Optional Marcar As Boolean)
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
Dim MyRecAux As New ADODB.Recordset
Dim cm As String

    frmMonitor.MostrarInfo CargaCadena(233)
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
        Do_Events
        If Modificaciones Then
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar where (tran_bar<>'*' OR tran_bar is null)")
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar")
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    
                    If !borrado = False Then
                        If !numero_eqp <> 0 Then
                            ' Equipo
                            CadenadeLog "Consulta Equipo " & CStr(!numero_eqp)
                            MyRecAux.Open "Select * from mbd.BarcodeEan13 where type=2 and number=" & CStr(!numero_eqp), MyConnObj, adOpenKeyset, adLockOptimistic
                        Else
                            ' Seccion
                            CadenadeLog "Consulta Seccion " & CStr(!secc_maqui)
                            MyRecAux.Open "Select * from mbd.BarcodeEan13 where type=1 and number=" & CStr(!secc_maqui), MyConnObj, adOpenKeyset, adLockOptimistic
                        End If
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            
                            If MyRecAux.EOF Then
                                MyRecAux.AddNew
'MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("netsales")))
'MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("Netsales")))
'MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetTotal")))
'MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetSS")))
'MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetPacket")))
'MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localsales")))
'MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localsales")))
'MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localTotal")))
'MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localSS")))
'MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localPacket")))

                                MyRecAux.Fields("netsales") = !CB_VNT
                                MyRecAux.Fields("nettotal") = !cb_mix
                                MyRecAux.Fields("netss") = !CB_EV1
                                MyRecAux.Fields("netpacket") = !CB_EV2
                                MyRecAux.Fields("localsales") = !CB_VNT
                                MyRecAux.Fields("localtotal") = !cb_mix
                                MyRecAux.Fields("localss") = !CB_EV1
                                MyRecAux.Fields("localpacket") = !CB_EV2
                                
                                CadenadeLog "Nuevo item"
                            End If
                            
                            If IsNull(!numero_eqp) Then
                                ' Seccion
                                MyRecAux.Fields("id") = 1000 + (!secc_maqui)
                                MyRecAux.Fields("number") = !secc_maqui
                                MyRecAux.Fields("type") = 1
                            Else
                                If !numero_eqp <> 0 Then
                                    ' Balanza
                                    MyRecAux.Fields("id") = 2000 + (!secc_maqui)
                                    MyRecAux.Fields("type") = 2
                                    MyRecAux.Fields("number") = !numero_eqp
                                Else
                                    ' Seccion
                                    MyRecAux.Fields("id") = 1000 + (!secc_maqui)
                                    MyRecAux.Fields("number") = !secc_maqui
                                    MyRecAux.Fields("type") = 1
                                End If
                            End If
                            If !secc_eqp = 0 Then
                                ' Red
                                MyRecAux.Fields("netsales") = !CB_VNT
                                MyRecAux.Fields("nettotal") = !cb_mix
                                MyRecAux.Fields("netss") = !CB_EV1
                                MyRecAux.Fields("netpacket") = !CB_EV2
                            Else
                                ' Local
                                MyRecAux.Fields("localsales") = !CB_VNT
                                MyRecAux.Fields("localtotal") = !cb_mix
                                MyRecAux.Fields("localss") = !CB_EV1
                                MyRecAux.Fields("localpacket") = !CB_EV2
                            End If
                               
                            
                            MyRecAux.Update
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            If IsNull(!numero_eqp) Then
                                frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                            Else
                                If !numero_eqp Then
                                    frmMonitor.MostrarDato CargaCadena(23) & " EQP.:" & Format(!numero_eqp, "000") & " --> OK"
                                Else
                                    frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                                End If
                            End If
                            .Edit
                            .Fields("tran_bar") = "*"
                            .Update

                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        If !numero_eqp <> 0 Then
                            cm = "delete from mbd.BarcodeEan13 where type=2 and number=" & CStr(!numero_eqp)
                        Else
                            cm = "delete from mbd.BarcodeEan13 where numbertype=1 and number=" & CStr(!secc_maqui)
                        End If
                        
                        MyConnObj.Execute cm
                        ''''''''''''''''''''
                        'frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tran_bar") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirBarras
                    
                Loop
            End If
        End With
    
SalirBarras:
    MyRecVar.Close
    Set MyRecVar = Nothing
salir:
    Mybase.Close
    Set Mybase = Nothing

End Sub


Public Sub MUE_EnviaVendedor(NSistema As Integer, Modificaciones As Boolean)
    '///////////////////////////////////////
    '// Envío de vendedores (TouchScale)    //
    '///////////////////////////////////////
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim cm As String
Dim lEsAlta As Boolean
Dim sColorTmp As String
Dim nColor As Integer
Dim MisEqp As dao.Recordset

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINVENDEDOR
    End If
    On Error GoTo 0
    On Error GoTo FINVENDEDOR
    
    Set MisEqp = Mybase.OpenRecordset("select * from equipos order by modelo asc")
    If MisEqp.EOF Then
        GoTo FINVENDEDOR
    Else
        MisEqp.MoveFirst
        Do While Not MisEqp.EOF
            If MisEqp.Fields("modelo") = 1 Then
                frmControl.Switch_mode 0
            Else
                frmControl.Switch_M MisEqp.Fields("prog_eqp")
            End If

            If EstadoBal = 0 Then
    
                Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where borrado=true")
                If Not MyRecVar.EOF Then
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = MyRecVar.Recordcount
                    MyRecVar.MoveFirst
                    Do While Not MyRecVar.EOF
                        If MisEqp.Fields("modelo") = 1 Then
                            CadenadeLog "Baja Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
                            cm = "delete from mbd.Vendors where id=" & CStr(MyRecVar.Fields("ident_vend"))
                            MyConnObj.Execute cm
                        End If
                        cm = "delete from mbd.localVendorKeys where vendorid=" & CStr(MyRecVar.Fields("ident_vend"))
                        MyConnObj.Execute cm
                        cm = "delete from mbd.localVendorPresel where vendorid=" & CStr(MyRecVar.Fields("ident_vend"))
                        MyConnObj.Execute cm
                        If Not MyRecVar.EOF Then MyRecVar.Movenext
                    Loop
                End If
                MyRecVar.Close
                Set MyRecVar = Nothing
           End If
           
           MisEqp.Movenext
            
        Loop
        'MisEqp.Close
        'Set MisEqp = Nothing
        cm = "delete from vendedor where borrado=true"
        Mybase.Execute cm
    End If
    
    
    MisEqp.MoveFirst
    Do While Not MisEqp.EOF
        If MisEqp.Fields("modelo") = 1 Then
            frmControl.Switch_mode 0
        Else
            frmControl.Switch_M MisEqp.Fields("prog_eqp")
        End If

        If EstadoBal = 0 Then
    
            If Modificaciones Then
                Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where tran_vend<>'*'")
            Else
                Set MyRecVar = Mybase.OpenRecordset("select * from vendedor")
            End If
            If Not MyRecVar.EOF Then
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                MyRecVar.MoveFirst
                Do While Not MyRecVar.EOF
                    If MisEqp.Fields("modelo") = 1 Then
                        frmMonitor.MostrarInfo CargaCadena(511) & " " & CStr(MyRecVar.Fields("ident_vend"))
                        On Error Resume Next
                        MyRecDes.Open "select * from mbd.Vendors where id=" & CStr(MyRecVar.Fields("ident_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            Call frmControl.Balanza_ComError(0)
                            GoTo FINVENDEDOR
                        End If
                        On Error GoTo FINVENDEDOR
                        If MyRecDes.EOF Then
                            MyRecDes.AddNew
                        End If
                        MyRecDes.Fields("id") = MyRecVar.Fields("ident_vend")
                        MyRecDes.Fields("name") = Trim(MyRecVar.Fields("nombre"))
                        MyRecDes.Update
                        MyRecDes.Close
                        Set MyRecDes = Nothing
                    End If
                    'mbd.localVendorKeys where vendorid=" & CStr(MyRecVar.Fields("ident_vend"))
                    On Error Resume Next
                    MyRecDes.Open "select * from mbd.localVendorKeys where id=" & CStr(MyRecVar.Fields("ident_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Err.Number <> 0 Then
                        Call frmControl.Balanza_ComError(0)
                        GoTo FINVENDEDOR
                    End If
                    On Error GoTo FINVENDEDOR
                    If MyRecDes.EOF Then
                        MyRecDes.AddNew
                    End If
                    MyRecDes.Fields("id") = MyRecVar.Fields("ident_vend")
                    MyRecDes.Fields("vendorid") = MyRecVar.Fields("ident_vend")
'  `Counter` int(11) DEFAULT '0',
'  `VendorKey` int(11) DEFAULT '0',
                    MyRecDes.Fields("counter") = MyRecVar.Fields("secc_maqui")
                    MyRecDes.Fields("vendorid") = MyRecVar.Fields("ident_vend")
                    MyRecDes.Fields("vendorkey") = MyRecVar.Fields("tec_vend")
                    MyRecDes.Update
                    MyRecDes.Close
                    Set MyRecDes = Nothing
                    
                    
                    Do_Events
                    MyRecVar.Movenext
        
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirVend
        
        
                Loop
            
            
            End If

        End If
           
        MisEqp.Movenext
            
    Loop
    MisEqp.Close
    Set MisEqp = Nothing

    cm = "update vendedor set tran_vend='*'"
    Mybase.Execute cm
    cm = "update vendedor set prog_vend='*'"
    Mybase.Execute cm

                    'MyRecVar.Edit
                    'MyRecVar.Fields("tran_vend") = "*"
                    'MyRecVar.Fields("prog_vend") = "*"
                    
                    'MyRecVar.Update

SalirVend:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINVENDEDOR:
    CadenadeLog "Error en MUE_EnviaVendedor:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
       
End Sub

Public Function Capa1A_GA_Erase_Totals(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para TouchScale  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim cm As String
    Dim nf As Integer
    Dim nC As Long
    Dim ncont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
   
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0

    CadenadeLog "Capa1A_GA_Erase_totals"
'*******
'//Reinicio numeraciÃ³n
'SQL_SendMessage("update numberingtickets set resetnumber=1");
'    mysql_free_result(RespSQL); // Liberamos la variable pues no se necesita respuesta
'SQL_SendMessage("update erasetickets set dattim=now() where type='T'");
'    mysql_free_result(RespSQL); // Liberamos la variable pues no se necesita respuesta
'//////////////////////
'SQL_SendMessage("truncate numtickets");
'mysql_free_result(RespSQL);
'SQL_SendMessage("truncate localnumtickets");
'mysql_free_result(RespSQL);
        
    
    CadenadeLog "Iniciar Numeración si Procede..." 'Qué hace en SAM
    If No_Reinit = 0 Then 'Iniciar numeración tiques...
        
    End If
    CadenadeLog "Borrado Documentos y Lineas Documento..."
    On Error Resume Next
    cm = "delete from mbd.hTicket where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.localhTicket where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from mbd.PayTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from mbd.localPayTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    'BorrarInventarioFin 'BorrarInventario
    If BorrarInventarioFin Then  '1.9.1 j.mer.
        On Error Resume Next
        'inventory_weight, inventory_units
        cm = "delete from mbd.ItemStock"
        MyConnObj.Execute cm
        ''''''''''''''''''''''''''''''''
        On Error GoTo 0
    End If
    If SC10_BorrarTotal Then 'Borrado totales diarios
       
    End If
    If SC10_Borrar_Eti_dat Then 'Borrado total IVA
    
    End If
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_GA_Erase_Totals = Resp
End Function


'Cabeceras de tique:
'.- NetStat: ‘N’, ‘L’.
'    Tique en red (‘N’) o local (‘L’).
'    Los tiques de red copiados en local se guardan en la tabla LocalHTicktes con NetStat=’N’.
'.- TckType: ‘N’, ‘O’
'    Tique “Normal” o  “Order” (encargo).
'    Los tiques encargo sólo se marcan como tales, no se realiza ninguna gestión especial.
'.- Status: ‘O’, ‘P’, ‘C’, ’D’, ‘R’, ‘N’, ‘X’
'    Estado del tique: abierto (‘O’), en impresión (‘P’), cerrado (‘C’), cancelado (‘D’),
'    recargado en tique abierto (‘R’),  recargado en tique cerrado (‘N’), cancelado manualmente (‘X’)
'Líneas de tique:
'.- LineType: ‘1’, ‘0’
'    Línea “normal” (‘1’) o cancelada (‘0’).
'.- TTare: ‘ ‘ , ‘W’, ‘M’, ‘I’
'    Tipo de tara aplicado en la operación. Tara acumulativa/pesada (‘W‘), manual (‘M’),
'    programada en el artículo (‘I’), sin tara (‘ ‘)

'adaptar sam
Public Sub LeeTickets(Optional Tabla As Integer, Optional ByVal lFinDia As Boolean)
    
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    Dim MyRecImp As New ADODB.Recordset
    
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim MisEqp As dao.Recordset
    Dim nLR As Integer
    
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim ncont As Long
    Dim dLImporte As Double
    Dim nLocal As Integer
    
    'If Dir(App.Path & "\hydramulti.exe") <> "" Or Dir(App.Path & "\noslaves.txt") <> "" Then
    '    nLocal = 0
    'Else
        nLocal = 1
    'End If
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'Call capturaClientes '''
    
    On Error GoTo fin
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MisEqp = Mybase.OpenRecordset("select * from equipos where modelo=1 and borrado=false")
    If MisEqp.EOF Then
        GoTo SalirLeeTique
    Else
        MisEqp.MoveFirst
        Do While Not MisEqp.EOF
            If MisEqp.Fields("modelo") = 1 Then
                frmControl.Switch_mode 0
            Else
                frmControl.Switch_M MisEqp.Fields("prog_eqp")
            End If

            If EstadoBal = 0 Then
            
            For nLR = 0 To 1 ' Red/Local
            
            'If MisEqp.Fields("modelo") <> 1 And nLR = 0 Then GoTo SigueNlr
            
            'CadenadeLog "paso 4"
            'status=4 --> Abierto.
            'status=1 --> Recargado == Cancelado.
            
            If nLR = 0 Then
                'If lFinDia And (Tabla = 1 Or Tabla = 2) Then
                '    cm = "update mbd.hTicket set status=0,enddate='" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' where ttype=" & CStr(Tabla) & " and status=4 and enddate='0000-00-00 00:00:00'"
                '    MyConnObj.Execute cm
                'End If
                MyRecSetTmp.Open "select * from mbd.hTickets where netstat='N' and (isnull(captured) or captured<>'*') and status<>'O' and status<>'P' and status<>'E' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
            Else
                'If lFinDia And (Tabla = 1 Or Tabla = 2) Then
                '    cm = "update mbd.localhTicket set status=0,enddate='" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' where ttype=" & CStr(Tabla) & " and status=4 and enddate='0000-00-00 00:00:00'"
                '    MyConnObj.Execute cm
                'End If
                MyRecSetTmp.Open "select * from mbd.hTickets where netstat='L' and (isnull(captured) or captured<>'*') and status<>'O' and status<>'P' and status<>'E' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
            End If
            'CadenadeLog "paso 5"
            If Not MyRecSetTmp.EOF Then
                Do While Not MyRecSetTmp.EOF
                    'CadenadeLog "paso 6"
                    Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
                    Fhora = Mid(Ffecha, 12, 8)
                    Ffecha = Mid(Ffecha, 1, 10)
                    'CadenadeLog "paso 7"
                    IdDoc = MyRecSetTmp.Fields("id")
                    'CadenadeLog "paso 8"
                    ''''medio de pago
                    '1=efectivo
                    '2=t.crédito
                    '3=cheque
                    Cabecera.tipoPago = 0
                    Cabecera.cantidadEntre = 0
                    
'                    If IsNull(Tabla) Or Tabla = 0 Then
''Select Case !TYPTIC --> Cabecera.tipoTiquet
''    Case 1
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
''    Case 2
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
''    Case 3
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
''    Case 4
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
''    Case 5
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
''    Case 6
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
''    Case 7
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
''    Case 8
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
''    Case 12
''        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Descart."
''End Select
'
'                            MyRecAux.Open "select * from mbd.PayTickets where idhTicket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
'                            'CadenadeLog "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc)
'                            If Not MyRecAux.EOF Then
'                                'CadenadeLog "paso 9"
'                                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
'                                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
'                                'CadenadeLog "paso 10"
'                            End If
'                            MyRecAux.Close
'                            Set MyRecAux = Nothing
'
'                    End If

                    'Impuestos
                    Cabecera.base0 = 0
                    Cabecera.por0 = 0
                    Cabecera.imp0 = 0
                    Cabecera.base1 = 0
                    Cabecera.por1 = 0
                    Cabecera.imp1 = 0
                    Cabecera.base2 = 0
                    Cabecera.por2 = 0
                    Cabecera.imp2 = 0
                    Cabecera.base3 = 0
                    Cabecera.por3 = 0
                    Cabecera.imp3 = 0
                    Cabecera.base4 = 0
                    Cabecera.por4 = 0
                    Cabecera.imp4 = 0
                       
                    'If nLR = 0 Then
                    '    'MyRecAux.Open "select COUNT(*) as lineas from mbd.lTicket where IdTicket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                    '    'CadenadeLog "paso 13"
                    '    nNumLin = MyRecSetTmp.Fields("nLines")
                    '    'MyRecAux.Close
                    '    'Set MyRecAux = Nothing
                    '    ReDim Cabecera.Lines(nNumLin)
                    'Else
                        'MyRecAux.Open "select COUNT(*) as lineas from mbd.locallTicket where Idticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        nNumLin = MyRecSetTmp.Fields("numlines")
                        'CadenadeLog "paso 16"
                        'MyRecAux.Close
                        'Set MyRecAux = Nothing
                        ReDim Cabecera.Lines(nNumLin)
                    'End If
                    
                    'recoger datos
                    Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
                    Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
                    Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
                    'CadenadeLog "paso 20"
                    'If IsNull(Tabla) Or Tabla = 0 Then
                        If IsNull(MyRecSetTmp.Fields("invoicenumber")) Or Len(MyRecSetTmp.Fields("invoicenumber")) = 0 Then
                            Cabecera.factura = Space(15)
                        Else
                            Cabecera.factura = Trim(MyRecSetTmp.Fields("invoicenumber"))
                            'CadenadeLog "paso 21"
                        End If
                        If IsNull(MyRecSetTmp.Fields("bcnumber")) Then
                            Cabecera.ean13 = ""
                        Else
                            Cabecera.ean13 = Trim(MyRecSetTmp.Fields("bcnumber"))
                        End If
                        'CadenadeLog "paso 22"
                        'If MyRecSetTmp.Fields("customer") <> "" Then
                            If IsNull(MyRecSetTmp.Fields("customer")) Then
                                Cabecera.NCliente = 0
                            Else
                                On Error Resume Next
                                Cabecera.NCliente = MyRecSetTmp.Fields("customer")
                                If Err.Number <> 0 Then
                                    On Error GoTo 0
                                    Cabecera.NCliente = 0
                                End If
                                On Error GoTo fin
                            End If
                        'Else
                        '    Cabecera.NCliente = 0
                        'End If
                        'CadenadeLog "paso 23"
                    'Else
                    '    Cabecera.factura = Space(15)
                    '    Cabecera.NCliente = 0
                    'End If
                    'CadenadeLog "paso 24"
                    
                    Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
                    Cabecera.hora = Fhora
                    Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
                    'If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("netamount")
                    'Else
                    '    Cabecera.ImporteTotal = MyRecSetTmp.Fields("amount")
                    'End If
                    'CadenadeLog "paso 25"
                    'If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.nLineas = nNumLin
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = nNumLin
                        Cabecera.LineasMensaje = nNumLin
                    'Else
                    '    Cabecera.nLineas = 1
                    '    Cabecera.LineaInicial = 1
                    '    Cabecera.LineaFinal = 1
                    '    Cabecera.LineasMensaje = 1
                    'End If
                    'cabecera.codigo ???
                    
                    Cabecera.EstadoTiquet = 16
                    'CadenadeLog "paso 26"
                    'If IsNull(Tabla) Or Tabla = 0 Then
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If nLR = 1 Then
                            Cabecera.CadenaEstadoTiquet = "Local"
                            Cabecera.tipoTiquet = 4 + MyRecSetTmp.Fields("workmode")
                        Else
                            Cabecera.CadenaEstadoTiquet = "Red"
                            Cabecera.tipoTiquet = 1 + MyRecSetTmp.Fields("workmode")
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'CadenadeLog "paso 27"
                        'If nLR = 0 Then
                        '    'CadenadeLog "paso 28"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 29"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    'CadenadeLog "paso 30"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 31"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                        'CadenadeLog "paso 32"
                        Cabecera.Trainning = False
                        'If MyRecSetTmp.Fields("tcktype") = "U" Then
                        '    Cabecera.tipoTiquet = 12
                        '    Cabecera.Trainning = True
                        'End If
                        'If MyRecSetTmp.Fields("status") = "R" Then
                        '    Cabecera.tipoTiquet = 13
                        '    Cabecera.EstadoTiquet = 64
                        'End If
                        'CadenadeLog "paso 33"
                        If MyRecSetTmp.Fields("status") = "X" Then 'cancelado
                            Cabecera.tipoTiquet = 14
                            Cabecera.EstadoTiquet = 64
                        End If
                        'CadenadeLog "paso 34"
                        'Cabecera.ImporteBruto = MyRecSetTmp.Fields("grossamount")
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Round(Cabecera.ImporteBruto - Cabecera.ImporteTotal, 2) 'más abajo se hace como porcentaje...
                        Cabecera.documentoabono = 0
                    
                    'Else
                    '    'CadenadeLog "paso 35"
                    '    If nLR = 0 Then
                    '        Cabecera.CadenaEstadoTiquet = "Red"
                    '    Else
                    '        Cabecera.CadenaEstadoTiquet = "Local"
                    '    End If
                    '
                    '    ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                    '    '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                    '    If Tabla = 1 Then
                    '        Cabecera.tipoTiquet = 2
                    '    Else
                    '        Cabecera.tipoTiquet = 3
                    '    End If
                    '
                    '    'Cabecera.CadenaEstadoTiquet = "Red"
                    '
                    '    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '
                    '    'Cabecera.ImporteBruto = dLImporte
                    '    Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                    '    Cabecera.Impuestos = 0
                    '    'Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal
                    '    Cabecera.documentoabono = 0
                    '    'Cabecera.Cambio = 0
                    '    'Cabecera.descuento = 0
                    '    'Cabecera.LineasCanceladas = 0
                    'End If
                    'CadenadeLog "paso 36"
                    If nLR = 0 Then
                        If MyRecSetTmp.Fields("workmode") = 0 Then
                            frmMonitor.MostrarDato CargaCadena(868) & " " & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket:" & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If MyRecSetTmp.Fields("workmode") = 1 Then
                            frmMonitor.MostrarDato CargaCadena(1412) & " " & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado:" & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            frmMonitor.MostrarDato CargaCadena(1413) & " " & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio:" & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    Else
                        If MyRecSetTmp.Fields("workmode") = 0 Then
                            frmMonitor.MostrarDato CargaCadena(1414) & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If MyRecSetTmp.Fields("workmode") = 1 Then
                            frmMonitor.MostrarDato CargaCadena(1415) & " " & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado Local:" & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            frmMonitor.MostrarDato CargaCadena(1416) & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio Local:" & Cabecera.Ntiquet & " Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    
                    End If
                    'CadenadeLog "paso 37"
                    'If Tabla = 0 Or IsNull(Tabla) Then
                        'If nLR = 0 Then
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '
                        'End If
                        'CadenadeLog "paso 38"
                        'CadenadeLog "7"
                        
                        Cabecera.LineasCanceladas = MyRecSetTmp.Fields("canceledlines")
                        
                        'también se vé afectado si es cancelado o no pero de mamento cerrado...
                        'Cabecera.EstadoTiquet = 16
                        '********************
                        ' este dato se introduce
                        ' tras pasar por actualiza_tabla_cabecera
                        'Incluir_en_Base As Boolean
                        'Cabecera.tipoPago = myrecsettmp.Fields("idmoneda")
                        'Cabecera.cantidadEntre = myrecsettmp.Fields("totalneto") + myrecsettmp.Fields("importecambio")
                        'If Cabecera.cantidadEntre <> 0 Then
                        '    Cabecera.Cambio = Round(Cabecera.ImporteTotal - Cabecera.cantidadEntre, 2) ' myrecsettmp.Fields("importecambio")
                        'Else
                            Cabecera.Cambio = 0
                        'End If
                        'Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")
                        'CadenadeLog "paso 39"
                        'lote As String
                        'lEncargo As Boolean
            
                        'LINEAS
            
                        nNumLin = 0
                        'If nLR = 0 Then
'Id, IdHTicket, NetStat, lineType, lineNumber, SaleForm, SaleFormEx, Item, NumPlu, Name, Counter, Family, Scale, Vendor, VendorName, Customer, CustomerName, lDatTim, PriceChanged, VATCode, VATPercent, VATValue, Weight, DrainedWeight, Tare, TTare, lineDiscount, Price, NetAmount, GrossAmount, Points, BovineCard, TraceCard, EarTag, BovineLot, ExpiryDate, BestBefDate, RemoveDate, FreezingDate, TicketDiscount, EAN13, EAN14, TimesPrinted, PrintScale, PrintDateTime, PrintTicketNumber, PrinttckType, ReloadScale, ReloadDateTime, ReloadTicketNumber, ReloadtckType, ReloadTo, CheckSu
                            MyRecAux.Open "select * from mbd.lTickets where  IdHTicket=" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'Else
                        '    MyRecAux.Open "select * from mbd.locallTicket where idticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'End If
                        'CadenadeLog "8"
                        'CadenadeLog "paso 40"
                        If Not MyRecAux.EOF Then
                            
                            'CadenadeLog "8.1"
                            
                            MyRecAux.MoveFirst
                            
                            'CadenadeLog "8.2"
                            
                            Do While Not MyRecAux.EOF
                                
                                'CadenadeLog "paso 41"
                                nNumLin = nNumLin + 1
                                If MyRecAux.Fields("saleform") = 0 Then
                                    Cabecera.Lines(nNumLin).EsPesado = True
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                Else
                                    Cabecera.Lines(nNumLin).EsPesado = False
                                    Cabecera.Lines(nNumLin).cantidad = Int(MyRecAux.Fields("weight")) ' / 1000)
                                End If
                                'CadenadeLog "paso 42"
                                'CadenadeLog "8.4"
                                'If Not IsNull(MyRecAux.Fields("numplu")) Then
                                'Cabecera.Lines(nNumLin).Plu = MyRecAux.Fields("numplu")
                                'Else
                                'Cabecera.Lines(nNumLin).Plu = 0
                                'End If
                                'Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                                Cabecera.Lines(nNumLin).desc = MyRecAux.Fields("name")
                                Cabecera.Lines(nNumLin).familia = MyRecAux.Fields("family")
                                Cabecera.Lines(nNumLin).FactorPeso = 1
                                Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                                'CadenadeLog "paso 43"
                                'CadenadeLog "8.5"
                                
                                Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                                Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                                
                                'CadenadeLog "paso 44"
                                'ver qué se hace IVA As Integer
                                Cabecera.Lines(nNumLin).IVA = MyRecAux.Fields("vatcode")
                                Cabecera.Lines(nNumLin).porcentaje = MyRecAux.Fields("vatpercent")
                                'CadenadeLog "paso 45"
                                'CadenadeLog "8.6"
                                
                                'If Not IsNull(MyRecAux.Fields("linediscount")) Then
                                'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("ticketdiscount") + MyRecAux.Fields("linediscount") 'porcentaje debido a la cabecera + porcentaje en linea
                                'Else
                                'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("ticketdiscount") + 0 'porcentaje debido a la cabecera + porcentaje en lineaCabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("ticketdiscount") + 0 'porcentaje debido a la cabecera + porcentaje en linea
                                'End If
                                'CadenadeLog "paso 46"
                                'If Not IsNull(MyRecAux.Fields("grossamount")) Then
                                'Cabecera.Lines(nNumLin).ImporteBruto = MyRecAux.Fields("grossamount") 'Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2) 'MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                                'Else
                                'Cabecera.Lines(nNumLin).ImporteBruto = 0
                                'End If
                                'CadenadeLog "paso 47"
                                Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("netamount") 'Round(MyRecAux.Fields("netamount") - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                'Cabecera.Lines(nNumLin).ImporteLinea = Round(Cabecera.Lines(nNumLin).ImporteLinea - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                
                                'CadenadeLog "paso 48"
                                'CadenadeLog "8.7"
                                
                                'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                                If MyRecAux.Fields("linetype") = 1 Then
                                    Cabecera.Lines(nNumLin).LineaCancelada = True
                                Else
                                    Cabecera.Lines(nNumLin).LineaCancelada = False
                                End If
                                'CadenadeLog "paso 49"
                                'ver qué se hace con TipoLinea As Byte
                                'STV (positivo o negativo)
                                'Linea(Bucle).ModificadoTemporal = False
                                If MyRecAux.Fields("netamount") > 0 Then
                                    'Cabecera.Lines(nNumLin).positive = 1
                                     Cabecera.Lines(nNumLin).FactorPeso = 1
                                     Cabecera.Lines(nNumLin).TipoLinea = 0
                                
                                Else
                                    'Cabecera.Lines(nNumLin).positive = 0
                                     Cabecera.Lines(nNumLin).FactorPeso = -1
                                     Cabecera.Lines(nNumLin).TipoLinea = 2
                                    
                                End If
                                'CadenadeLog "paso 50"
            '                    If Cabecera.Lines(nNumLin).positive = 1 Then
            '                         Cabecera.Lines(nNumLin).FactorPeso = 1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 0
            '                    Else
            '                         Cabecera.Lines(nNumLin).FactorPeso = -1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 2
            '                    End If
                                'STP (pesado o no pesado)
                                If Cabecera.Lines(nNumLin).EsPesado Then
                                     'Cabecera.Lines(nNumLin).Weighted = 1
                                Else
                                     'Cabecera.Lines(nNumLin).Weighted = 0
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                                End If
                                
                                'CadenadeLog "paso 51"
                                ' STC (cancelado o no cancelado)
                                'If IsNull() Then
                                '
                                'Else
                                    If Cabecera.Lines(nNumLin).LineaCancelada Then
                                         Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                                    End If
                                'End If
                                'CadenadeLog "paso 52"
                                'If MyRecAux.Fields("pricechanged") <> 0 Then
                                '    Cabecera.Lines(nNumLin).SehaModificado = True
                                'Else
                                '    Cabecera.Lines(nNumLin).SehaModificado = False
                                'End If
                                '¿? ModificadoTemporal As Boolean
                                'CadenadeLog "paso 53"
                                'CadenadeLog "8.8"
                                'c2f 1.1.1 aclarar ficha Vacuno...
                                'If IsNull(MyRecAux.Fields("eartag")) Or MyRecAux.Fields("eartag") = "" Then
                                    
                                    'If Not IsNull(MyRecAux.Fields("bovinecard")) Then
                                    '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("bovinecard")
                                    'Else
                                        Cabecera.Lines(nNumLin).FichaVacuno = 0
                                    'End If
                                
                                'Else
                                '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("eartag")
                                'End If
                                'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                                'Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'CadenadeLog "8.9"
                                
                                MyRecAux.Movenext
                            Loop
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            'CadenadeLog "paso 54"
                            
                            Actualiza_Tabla_Cabecera Mybase, Cabecera
                            
                            'CadenadeLog "paso 55"
                            Do_Events
                            If Cabecera.Incluir_en_Base = True Then
                                Actualiza_Tabla_Tickets Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtarti Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_GtVend Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtHora Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtsecs Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtInfVen Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Log Mybase, Cabecera
                            End If
                            
                            'CadenadeLog "paso 56"
                        End If
            
            
            
                        '''''''''''''''
                        'marcar...
                        'cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
                        'If nLR = 0 Then
                            cm = "update mbd.hTickets set Captured='*' where id=" & CStr(IdDoc)
                        'Else
                        '    cm = "update mbd.localhTicket set Captured=1 where id=" & CStr(IdDoc)
                        'End If
                        MyConnObj.Execute cm
                        
                        CadenadeLogTotales "Marcando:" & "update mbd.hTickets set Captured='*' where id=" & CStr(IdDoc)
                        
                        ''''''''''
                        'CadenadeLog "paso 57"
                    'Else 'REVISAR...
                    '
                    '
                    '        'CadenadeLog "paso 58"
                    '        nNumLin = 1
                    '        If MyRecSetTmp.Fields("saleform") = 1 Then
                    '            Cabecera.Lines(nNumLin).EsPesado = True
                    '            Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                    '        Else
                    '            Cabecera.Lines(nNumLin).EsPesado = False
                    '            Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                    '        End If
                    '        'CadenadeLog "paso 59"
                    '        'CadenadeLog "8.4"
                    '
                    '        Cabecera.Lines(nNumLin).FactorPeso = 1
                    '        Cabecera.Lines(nNumLin).precio = MyRecSetTmp.Fields("price")
                    '
                    '        'CadenadeLog "8.5"
                    '        'CadenadeLog "paso 60"
                    '        Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                    '        Cabecera.Lines(nNumLin).CodigoArticulo = MyRecSetTmp.Fields("item")
                    '        'ver qué se hace IVA As Integer
                    '        Cabecera.Lines(nNumLin).IVA = 0
                    '
                    '        'CadenadeLog "paso 61"
                    '        'CadenadeLog "8.6"
                    '
                    '        Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("linediscount") 'porcentaje
                    '        Cabecera.Lines(nNumLin).ImporteLinea = MyRecSetTmp.Fields("amount") 'MyRecSetTmp.Fields("importeconimpuestos") - MyRecSetTmp.Fields("dtoscabeceraconimpuestos")
                    '        Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecSetTmp.Fields("weight") * MyRecSetTmp.Fields("price"), 2) 'MyRecSetTmp.Fields("importe") - MyRecSetTmp.Fields("dtoscabecera")
                    '        'CadenadeLog "paso 62"
                    '        'CadenadeLog "8.7"
                    '
                    '        'If MyRecSetTmp.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                    '        If MyRecSetTmp.Fields("linetype") = 0 Then
                    '            Cabecera.Lines(nNumLin).LineaCancelada = True
                    '        Else
                    '            Cabecera.Lines(nNumLin).LineaCancelada = False
                    '        End If
                    '        'CadenadeLog "paso 63"
                    '        'ver qué se hace con TipoLinea As Byte
                    '        'STV (positivo o negativo)
                    '        'Linea(Bucle).ModificadoTemporal = False
                    '        If MyRecSetTmp.Fields("amount") > 0 Then
                    '            'Cabecera.Lines(nNumLin).positive = 1
                    '             Cabecera.Lines(nNumLin).FactorPeso = 1
                    '             Cabecera.Lines(nNumLin).TipoLinea = 0
                    '
                    '        Else
                    '            'Cabecera.Lines(nNumLin).positive = 0
                    '             Cabecera.Lines(nNumLin).FactorPeso = -1
                    '             Cabecera.Lines(nNumLin).TipoLinea = 2
                    '
                    '        End If
                    '        'CadenadeLog "paso 64"
        '           '         If Cabecera.Lines(nNumLin).positive = 1 Then
        '           '              Cabecera.Lines(nNumLin).FactorPeso = 1
        '           '              Cabecera.Lines(nNumLin).TipoLinea = 0
        '           '         Else
        '           '              Cabecera.Lines(nNumLin).FactorPeso = -1
        '           '              Cabecera.Lines(nNumLin).TipoLinea = 2
        '           '         End If
                    '        'STP (pesado o no pesado)
                    '        If Cabecera.Lines(nNumLin).EsPesado Then
                    '             'Cabecera.Lines(nNumLin).Weighted = 1
                    '        Else
                    '             'Cabecera.Lines(nNumLin).Weighted = 0
                    '             Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                    '        End If
                    '        'CadenadeLog "paso 65"
                    '
                    '        ' STC (cancelado o no cancelado)
                    '        If Cabecera.Lines(nNumLin).LineaCancelada Then
                    '             Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                    '        End If
                    '        'CadenadeLog "paso 66"
                    '        'If MyRecSetTmp.Fields("precio") <> MyRecSetTmp.Fields("preciodefecto") Then
                    '        '    Cabecera.Lines(nNumLin).SehaModificado = True
                    '        'Else
                    '            Cabecera.Lines(nNumLin).SehaModificado = False
                    '        'End If
                    '        '¿? ModificadoTemporal As Boolean
                    '        'CadenadeLog "paso 67"
                    '        'CadenadeLog "8.8"
                    '        'c2f 1.1.1 aclarar ficha Vacuno...
                    '        'If IsNull(MyRecSetTmp.Fields("eartag")) Or MyRecSetTmp.Fields("eartag") = "" Then
                    '            Cabecera.Lines(nNumLin).FichaVacuno = 0
                    '        'Else
                    '        '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecSetTmp.Fields("eartag")
                    '        'End If
                    '        'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                    '        'Cabecera.descuento = Cabecera.descuento + MyRecSetTmp.Fields("dtoslineaconimpuestos")
                    '
                    '        'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("dtoslineaconimpuestos")
                    '
                    '        'CadenadeLog "8.9"
                   '
                    '    'CadenadeLog "paso 68"
                    '    Actualiza_Tabla_Cabecera Mybase, Cabecera
                    '    Do_Events
                    '    If Cabecera.Incluir_en_Base = True Then
                    '        Actualiza_Tabla_Tickets Mybase, Cabecera
                    '        Do_Events
                    '        Actualiza_Tabla_Gtarti Mybase, Cabecera
                    '        Do_Events
                    '        Actualiza_Tabla_GtVend Mybase, Cabecera
                    '        Do_Events
                    '        actualiza_tabla_gtHora Mybase, Cabecera
                    '        Do_Events
                    '        Actualiza_Tabla_Gtsecs Mybase, Cabecera
                    '        Do_Events
                    '        actualiza_tabla_gtInfVen Mybase, Cabecera
                    '        Do_Events
                    '        Actualiza_Tabla_Log Mybase, Cabecera
                    '    End If
                    '    'CadenadeLog "paso 69"
                    '    'marcar...
                    '    If nLR = 0 Then
                    '        If Tabla = 1 Then
                    '            cm = "update pcscale.packetlabels set customer='9999999999' where id=" & CStr(IdDoc)
                    '        Else
                    '            cm = "update pcscale.sservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                    '        End If
                    '    Else
                    '        If Tabla = 1 Then
                    '            cm = "update pcscale.localpacketlabels set customer='9999999999' where id=" & CStr(IdDoc)
                    '        Else
                    '            cm = "update pcscale.localsservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                    '        End If
                    '    End If
                    '    MyConnObj.Execute cm
                    '    ''''''''''
                    '    'CadenadeLog "paso 70"
                    'End If
                    
                    MyRecSetTmp.Update
        
                    MyRecSetTmp.Movenext
                
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirLeeTique
                    Do_Events
                    
                Loop
            End If
            'CadenadeLog "paso 71"
            'On Error Resume Next
            If MyRecSetTmp.State <> 0 Then
                MyRecSetTmp.Close
                Set MyRecSetTmp = Nothing
            End If
            'On Error GoTo 0
SigueNlr:
            'CadenadeLog "paso 72"
            Next nLR
            
            End If
            'CadenadeLog "paso 73"
            MisEqp.Movenext
            
        Loop
        
        MisEqp.Close
        Set MisEqp = Nothing
        
    End If
    
SalirLeeTique:
''''''''
    On Error GoTo 0
    On Error GoTo fin
    'myRec.Close
    'Set myRec = Nothing
    If MyRecSetTmp.State <> 0 Then
        MyRecSetTmp.Close
        Set MyRecSetTmp = Nothing
    End If
    If Mybase.name <> "" Then
    Mybase.Close
    Set Mybase = Nothing
    End If
fin:
    If Err.Number <> 0 Then
        CadenadeLog "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
        On Error GoTo 0
        On Error GoTo FIN1
        If MyConnObj.State <> 0 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            Do_Events
            Sleep (200)
            frmControl.Switch_mode 0
            Do_Events
            Sleep (200)
            EstadoBal = 1000
        End If
    End If
FIN1:
    On Error GoTo 0

End Sub

'DROP TABLE IF EXISTS `mbd`.`Counters`;
'CREATE TABLE  `mbd`.`Counters` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(64) DEFAULT '',
'  `Icon` varchar(127) DEFAULT NULL,
'  `Turn` int(10) unsigned DEFAULT '1',
'  `Description` varchar(127) DEFAULT '',
'  `Text1` varchar(255) DEFAULT '',
'  `Text2` varchar(255) DEFAULT '',
'  `Text3` varchar(255) DEFAULT '',
'  `Text4` varchar(255) DEFAULT '',
'  `Text5` varchar(255) DEFAULT '',
'  `Text6` varchar(255) DEFAULT '',
'  `Text7` varchar(255) DEFAULT '',
'  `Text8` varchar(255) DEFAULT '',
'  `Text9` varchar(255) DEFAULT '',
'  `Text10` varchar(255) DEFAULT '',
'  PRIMARY KEY (`Code`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaDesSec()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nResp As Integer

'    nResp = ElPing(sAddIpTouch)
'    If nResp = 1 Then

    On Error GoTo FINDESSEC
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select * from seccion where borrado=false")
    On Error GoTo FINDESSEC
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            MyRecDes.Open "Select * from mbd.Counters where code=" & CStr(MyRecVar.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            frmMonitor.MostrarInfo CargaCadena(1417) & " " & CStr(MyRecVar.Fields("secc_maqui"))
            MyRecDes.Fields("code") = MyRecVar.Fields("secc_maqui")
            If Not IsNull(MyRecVar.Fields("nombre")) Then
                MyRecDes.Fields("name") = Trim(MyRecVar.Fields("nombre"))
            Else
                MyRecDes.Fields("name") = "Sección: " & CStr(MyRecVar.Fields("secc_maqui"))
            End If
            If Not IsNull(MyRecVar.Fields("descripcio")) Then
                MyRecDes.Fields("description") = Trim(MyRecVar.Fields("descripcio"))
            Else
                MyRecDes.Fields("description") = "---" & CargaCadena(47) & " " & CStr(MyRecVar.Fields("secc_maqui")) & "---"
            End If
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirDesSec
            
        Loop
    
    
    End If
SalirDesSec:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINDESSEC:
    CadenadeLog "Error en MUE_EnviaDesSec:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
'    Else
'        canCelar = 1
'    End If
End Sub


Public Sub MUE_EnviaCabLey()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCABLEY
    End If
    On Error GoTo 0
    
    Set MyRecVar = Mybase.OpenRecordset("select * from trade")
    If Not MyRecVar.EOF Then
            
            MyRecDes.Open "Select * from mbd.Texts", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                'Call frmControl.Balanza_ComError(0)
                GoTo FINCABLEY
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            
            frmMonitor.MostrarInfo CargaCadena(1418)
            MyRecDes.Fields("id") = 1
            MyRecDes.Fields("number") = 1
            MyRecDes.Fields("header1") = MyRecVar.Fields("company")
            MyRecDes.Fields("header2") = MyRecVar.Fields("name")
            MyRecDes.Fields("header3") = MyRecVar.Fields("address")
            MyRecDes.Fields("legend1") = MyRecVar.Fields("quarter")
            MyRecDes.Fields("legend2") = MyRecVar.Fields("web")
            MyRecDes.Fields("legend3") = MyRecVar.Fields("shopweb")
            MyRecDes.Fields("legend4") = MyRecVar.Fields("description")
            MyRecDes.Fields("legend5") = MyRecVar.Fields("slogan")
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCabLey
    
    End If
SalirCabLey:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCABLEY:
    CadenadeLog "Error en MUE_EnviaCabLey:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
End Sub

'***********************************
'////////////////////////////
'TouchScale --> BASE MAESTRA
Public Sub AN_BaseMaestra(Modificaciones As Boolean, Optional ByVal lC As Boolean, Optional ByVal nEnv As Integer)
    Dim bucle As Integer
    Dim Base As dao.Database
    Dim SeBorra As Boolean
    Dim Beco As Long
    Dim Registro As dao.Recordset
    Dim nEnviados As Integer
    
    On Error GoTo fin

    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = False

    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = True
    frmMonitor.Label1 = ""
    
    lCogeTiquet = False
    canCelar = 0
    CancelHumano = False
    
    MostrarMonitor
    
    MUE_LanzaBaseMaestra Modificaciones, nEnv
    nEnviados = 2

    '*****************************
    ' borrado de datos eliminados
    'SE REALIZARÁN EN ENVIO...
    '*****************************
    'Set Base = OpenDatabase(Base_General) '1.7.2 OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo where borrado=true and tran_plu='*' and tran_pluSC10='*' and tran_en='*' and tran_el='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If (Not .EOF) Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from publi where borrado=true and tran_pub='*' and tran_pubSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from cabley where borrado=true and tran_cab='*' and tran_cabSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from codbar where borrado=true and tran_bar='*' and tran_barSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from vendedor where borrado=true and tran_vend='*' and tran_vendSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from teclas where borrado=true")
    ''2.0.30 ("select * from teclas where borrado=true and tran_tec='*' and tran_tecSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
fin:
On Error GoTo 0

    On Error Resume Next
    Set Base = Nothing
    frmMonitor.Label1.Caption = CargaCadena(445)  '"Comunicación Finalizada"
    frmMonitor.cmdmonitorstop.Enabled = False
    On Error GoTo 0
    HabilitarSalir
    lCogeTiquet = True
End Sub

Public Function MUE_LanzaBaseMaestra(Modificaciones As Boolean, Optional ByVal nEnv As Integer) As Boolean
'************************************************************************************
' envío de base maestra, realizando un ciclo por todas los sistemas existentes
'************************************************************************************
    Dim bucle As Integer
    Dim Marcar As Boolean
    Dim ErrorPrevio As Boolean
    Dim cm As String
    
    ErrorPrevio = False
    Marcar = False
    
    MUE_Sistemas
    
    For bucle = 0 To SisEur.NSistemas - 1
        If MultiEuroscale Then
            frmControl.Switch_mode 700 + bucle
        Else
            frmControl.Switch_mode 0
        End If
        If CancelHumano Then Exit For
        canCelar = 0
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaDesSec
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        '1.1.3
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 7) Then MUE_EnviaArticuloNW bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaCabLey
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
'Mue_Env_Gen
#If BALSAM = True Then
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then Mue_Env_Gen bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If

#End If
        ''''''''''''''''''''''''''''''''''
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaVendedor bucle, Modificaciones
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 2) Then MUE_EnviaCLientes bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 4) Then MUE_EnviaCodigosBarras bucle, Modificaciones
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 5) Then MUE_EnviarFichasVacuno bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 Then Call capturaPaises
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
        frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
        If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
        On Error Resume Next
        
        MyConnObj.Close
        Set MyConnObj = Nothing
        On Error GoTo 0
        EstadoBal = 1000
    
    
    Next bucle
    On Error GoTo 0
    
End Function

Public Sub KillTECLAS()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset

    frmControl.Switch_mode 0
    If EstadoBal = 0 Then


        On Error Resume Next
        CadenadeLog "mbd.ItemKeys"
        cm = "delete from mbd.ItemKeys"
        MyConnObj.Execute cm
        If Err.Number = 0 Then
            MsgBox "TECLAS BORRADAS.", vbExclamation
        End If
        On Error GoTo 0
        
        CadenadeLog "BORRAR TECLAS..."


    End If
    
End Sub

Public Sub KillARTICLES()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
Dim a As Variant

    a = MsgBox("EL BORRADO DE ARTÍCULOS CONLLEVA LA CAPTURA Y BORRADO DE TOTALES, ¿DESEA CONTINUAR?", vbYesNo)
    If a = vbYes Then
        frmControl.Switch_mode 0
        If EstadoBal = 0 Then
    
            CancelHumano = False
            frmMonitor.LblError = ""
            frmMonitor.Show
            frmMonitor.SetFocus
            
            '''''''
            'Call capturaClientes
            ''''''''
            LeeTickets 0
            'LeeTickets 1, True
            'LeeTickets 2, True
            'Revisar SAM
            'If canCelar = 0 Then MUE_SolicitaPedido 0, False
            'If canCelar = 0 Then MUE_SolicitaInventario 0, False
            If GA_Reinit Then
               Capa1A_GA_Erase_Totals (0)
            Else
               Capa1A_GA_Erase_Totals (1)
            End If
            If canCelar = 0 Then
                On Error Resume Next
                CadenadeLog "mbd.ItemKeys"
                cm = "delete from mbd.ItemKeys"
                MyConnObj.Execute cm
                On Error GoTo 0
        
        
                On Error Resume Next
                cm = "delete from mbd.ItemRates where code<>0"
                MyConnObj.Execute cm
                CadenadeLog "mbd.Items"
                cm = "delete from  where mbd.Items code<>0"
                MyConnObj.Execute cm
                If Err.Number = 0 Then
                    MsgBox "ARTICULOS Y TECLAS BORRADOS.", vbExclamation
                End If
                On Error GoTo 0
                
                CadenadeLog "BORRAR ARTICULOS..."
                
                frmMonitor.Hide
                frmEpelsa.SetFocus
                
            End If
            
        End If
    End If
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////
'Adaptar SAM
Public Sub BorraSrv()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
On Error Resume Next
    
    CadenadeLog "BORRASRV..."
    
    On Error Resume Next
    CadenadeLog "mbd.BarcodeCode128"
    cm = "delete from mbd.BarcodeCode128"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.BarcodeEan128"
    cm = "delete from mbd.BarcodeEan128"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.BarcodeEan13"
    cm = "delete from mbd.BarcodeEan13"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.BarcodeEan14"
    cm = "delete from mbd.BarcodeEan14"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.BarcodeGS1"
    cm = "delete from mbd.BarcodeGS1"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Bovine"
    cm = "delete from mbd.Bovine"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "mbd.Card"
    cm = "delete from mbd.Card"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "mbd.Company"
    cm = "delete from mbd.Company"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "mbd.Counters"
    cm = "delete from mbd.Counters"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "mbd.Countries"
    cm = "delete from mbd.Countries"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "mbd.Customers"
    cm = "delete from mbd.Customers"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Families"
    cm = "delete from mbd.Families"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemFavs"
    cm = "delete from mbd.ItemFavs"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemKeys"
    cm = "delete from mbd.ItemKeys"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemOffers"
    cm = "delete from mbd.ItemOffers"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemRates"
    cm = "delete from mbd.ItemRates"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemStepPrice"
    cm = "delete from mbd.ItemStepPrice"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.ItemStock"
    cm = "delete from mbd.ItemStock"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Items"
    cm = "delete from mbd.Items"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Nutritional"
    cm = "delete from mbd.Nutritional"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.PayTicket"
    cm = "delete from mbd.PayTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localPayTicket"
    cm = "delete from mbd.localPayTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Texts"
    cm = "delete from mbd.Texts"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.VatTicket"
    cm = "delete from mbd.VatTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localVatTicket"
    cm = "delete from mbd.localVatTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.VendorLog"
    cm = "delete from mbd.VendorLog"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localVendorLog"
    cm = "delete from mbd.localVendorLog"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.VendorPresel"
    cm = "delete from mbd.VendorPresel"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localVendorPresel"
    cm = "delete from mbd.localVendorPresel"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.Vendors"
    cm = "delete from mbd.Vendors"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localVendorKeys"
    cm = "delete from mbd.localVendorKeys"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.hTicket"
    cm = "delete from mbd.hTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.localhTicket"
    cm = "delete from mbd.localhTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.lTicket"
    cm = "delete from mbd.lTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "mbd.locallTicket"
    cm = "delete from mbd.locallTicket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
    End If

On Error GoTo 0

End Sub

Public Function ElPing(ElHost As String) As Integer
    '///////////////////////////////////////////////
    '//  Lanza ping a la dirección que se indique //
    '//  llamando al OCX de ping                  //
    '///////////////////////////////////////////////
    Dim Contador As Integer
    Dim Echo As ICMP_ECHO_REPLY
    Dim Resp As Long
    If Dir(App.Path & "\noping.txt") <> "" Then
        ElPing = 1
    Else
        Resp = ICMP_Ping(ElHost, "alive", Echo)
        If Resp = 0 Then
            ElPing = 1
        Else
            ElPing = 0
        End If
    End If
End Function

'''''''''''''''''''''
'''STOCK + INVENTARIO
'''''''''''''''''''''
'DROP TABLE IF EXISTS `pcscale`.`stock`;
'CREATE TABLE  `pcscale`.`stock` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Inventory_Weight` decimal(12,3) DEFAULT '0.000',
'  `Inventory_Units` int(11) DEFAULT '0',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `stock_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
''''''''''''''''''''''
'''PEDIDO
''''''''''''''''''''''
'DROP TABLE IF EXISTS `pcscale`.`shoporder`;
'CREATE TABLE  `pcscale`.`shoporder` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Observations` varchar(127) DEFAULT NULL,
'  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `shoporder_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;

Public Sub MUE_SolicitaPedido(NSistema As Integer, FinalDia As Boolean)
    Dim RutaPedido As String
    Dim RutaStock As String
    'Dim Mybase As DAO.Database
    'Dim MyRecVar As DAO.Recordset
    Dim nf As Integer
    'Dim nG As Integer
    Dim sF As String
    'Dim sG As String
    On Error GoTo PedFin
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        If Not HayMulti Then
            RutaPedido = App.Path & "\01ped.sdf"
            'RutaPedido = App.Path & "\01sto.sdf"
        Else
            If TiendaActual < 100 Then
                RutaPedido = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "ped.sdf" '& Format(nsistema, "00")
                'RutaStock = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "sto.sdf." & Format(nsistema, "00")
            Else
                RutaPedido = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "ped.sdf" '& Format(nsistema, "000")
                'RutaStock = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "sto.sdf." & Format(nsistema, "000")
            End If
        End If
        frmMonitor.MostrarInfo CargaCadena(1106)
        'If Dir(App.Path & "\pedplusec.txt") <> "" Then
        '    Resp = frmControl.Balanza.Query_Order(RutaPedido, TiendaActual)
        'Else
        '    Resp = frmControl.Balanza.Query_Order(RutaPedido)
        'End If
        'Set Mybase = DAO.OpenDatabase(Base_General)
        'MyRecSet.CursorLocation = adUseClient
        
        MyRecSet.Open "Select * from pcscale.shoporder where code<>0", MyConnObj, adOpenKeyset, adLockOptimistic '1.8.9 instr(lcase(pcscale.shoporder.observations),'(captured)')=0", MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecSet.EOF Then
            nf = FreeFile()
            '1.8.9 Open RutaPedido For Append As #nf
            Open RutaPedido For Output As #nf
            'nG = FreeFile()
            'Open RutaStock For Output As #nG
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                sF = Format(MyRecSet.Fields("code"), "000000")
                If MyRecSet.Fields("saleform") <> 0 Then
                    sF = sF & "W"
                    sF = sF & Format(MyRecSet.Fields("weight") * 1000, "00000000")
                Else
                    sF = sF & "U"
                    sF = sF & Format(MyRecSet.Fields("units"), "00000000")
                End If
                'sG = sF
                sF = sF & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                If Len(sF) < 40 Then
                    sF = sF & Mid("                                        ", 1, 40 - Len(sF))
                End If
                sF = sF & Format(MyRecSet.Fields("dattim"), "dd-mm-yy hh:mm:ss")
                Print #nf, sF
                '1.8.9 MyRecSet.Fields("observations") = Mid(MyRecSet.Fields("observations"), 1, 110) & "(captured)"
                '1.8.9 MyRecSet.Update
                MyRecSet.Movenext
            Loop
            Close #nf
        
        End If
        MyRecSet.Close
        Set MyRecSet = Nothing
        
        
        
        'If canCelar = 0 Then
        '    If (FinalDia And BorrarPedidoFin) Or BorrarPedido Then
        '        frmMonitor.MostrarInfo CargaCadena(1108)
        '        frmControl.Balanza.Erase_Order
        '    End If
        '
        'End If
        frmMonitor.Label1.Caption = CargaCadena(445)
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = True
    End If
    If canCelar = 0 Then
        frmMonitor.MostrarDato CargaCadena(1096) & " " & RutaPedido
    Else
        On Error Resume Next
        'Kill RutaPedido
    End If
PedFin:
End Sub

Public Sub MUE_SolicitaInventario(NSistema As Integer, FinalDia As Boolean)
    Dim RutaInventario As String
    Dim RutaStock As String
    Dim nf As Integer
    Dim nG As Integer
    Dim sF As String
    Dim sG As String
    Dim cm As String
    
    On Error GoTo PedFin
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        If Not HayMulti Then
            RutaInventario = App.Path & "\01inv.sdf"
            RutaStock = App.Path & "\01sto.sdf"
        Else
            If TiendaActual < 100 Then
                RutaInventario = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "inv.sdf" '& Format(nsistema, "00")
                RutaStock = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "sto.sdf" '& Format(nsistema, "00")
            Else
                RutaInventario = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "inv.sdf" '& Format(nsistema, "000")
                RutaStock = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "sto.sdf" '& Format(nsistema, "000")
            End If
        End If
        frmMonitor.MostrarInfo "Inventory & Stock."
'DROP TABLE IF EXISTS `pcscale`.`stock`;
'CREATE TABLE  `pcscale`.`stock` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Inventory_Weight` decimal(12,3) DEFAULT '0.000',
'  `Inventory_Units` int(11) DEFAULT '0',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `stock_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
        MyRecSet.Open "Select * from pcscale.stock where code<>0", MyConnObj, adOpenKeyset, adLockOptimistic '1.8.9 instr(lcase(pcscale.stock.name),'(captured)')=0", MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecSet.EOF Then
            nf = FreeFile()
            '1.8.9 Open RutaInventario For Append As #nf
            Open RutaInventario For Output As #nf
            nG = FreeFile()
            '1.8.9 Open RutaStock For Append As #nG
            Open RutaStock For Output As #nG
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                sF = Format(MyRecSet.Fields("code"), "000000")
                sG = sF
                If MyRecSet.Fields("saleform") <> 0 Then
                    sG = sG & "W"
                    sG = sG & Format(MyRecSet.Fields("weight") * 1000, "00000000")
                    sF = sF & "W"
                    sF = sF & Format(MyRecSet.Fields("inventory_weight") * 1000, "00000000")
                Else
                    sG = sG & "U"
                    sG = sG & Format(MyRecSet.Fields("units"), "00000000")
                    sF = sF & "U"
                    sF = sF & Format(MyRecSet.Fields("inventory_units"), "00000000") 'inventory_weight, inventory_units
                End If
                
                sF = sF & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                sG = sG & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                
                If Len(sF) < 40 Then
                    sF = sF & Mid("                                        ", 1, 40 - Len(sF))
                End If
                If Len(sG) < 40 Then
                    sG = sG & Mid("                                        ", 1, 40 - Len(sG))
                End If
                
                sF = sF & Format(Now, "dd-mm-yy hh:mm:ss")
                sG = sG & Format(Now, "dd-mm-yy hh:mm:ss")
                
                Print #nf, sF
                Print #nG, sG
                
                'MyRecSet.Fields("name") = Mid(MyRecSet.Fields("name"), 1, 110) & "(captured)"
                
                'MyRecSet.Update
                MyRecSet.Movenext
                
            Loop
            Close #nf
            Close #nG
        End If
        MyRecSet.Close
        Set MyRecSet = Nothing
        
        frmMonitor.Label1.Caption = CargaCadena(445)
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = True
    
    End If
    If canCelar = 0 Then
        frmMonitor.MostrarDato "Inventory:" & " " & RutaInventario
        frmMonitor.MostrarDato "Stock:" & " " & RutaStock
        If BorrarInventario Then '1.9.1 j.mer.
            On Error GoTo 0
            On Error Resume Next
            'inventory_weight, inventory_units
            cm = "update pcscale.stock set inventory_weight=0 where code<>0"
            MyConnObj.Execute cm
            ''''''''''''''''''''''''''''''''
            On Error GoTo 0
            On Error Resume Next
            'inventory_weight, inventory_units
            cm = "update pcscale.stock set inventory_units=0 where code<>0"
            MyConnObj.Execute cm
            ''''''''''''''''''''''''''''''''
            On Error GoTo 0
            On Error Resume Next
        End If
    Else
        On Error Resume Next
    End If
PedFin:
End Sub

Public Sub carga_data_Vitamin()
Dim nf As Integer
Dim nC As Integer
Dim sC As String
    If Dir(App.Path & "\vitaname.txt") = "" Then
        MyVM.data(0).name = "Vitamina A"
        MyVM.data(1).name = "Vitamina D"
        MyVM.data(2).name = "Vitamina E"
        MyVM.data(3).name = "Vitamina K"
        MyVM.data(4).name = "Vitamina C"
        MyVM.data(5).name = "Tiamina"
        MyVM.data(6).name = "Riboflavina"
        MyVM.data(7).name = "Niacina"
        MyVM.data(8).name = "Vitamina B6"
        MyVM.data(9).name = "Ácido fólico"
        MyVM.data(10).name = "Vitamina B12"
        MyVM.data(11).name = "Biotina"
        MyVM.data(12).name = "Ácido pantoténico"
        MyVM.data(13).name = "Potasio"
        MyVM.data(14).name = "Cloro"
        MyVM.data(15).name = "Calcio"
        MyVM.data(16).name = "Fósforo"
        MyVM.data(17).name = "Magnesio"
        MyVM.data(18).name = "Hierro"
        MyVM.data(19).name = "Zinc"
        MyVM.data(20).name = "Cobre"
        MyVM.data(21).name = "Manganeso"
        MyVM.data(22).name = "Flúor"
        MyVM.data(23).name = "Selenio"
        MyVM.data(24).name = "Cromo"
        MyVM.data(25).name = "Molibdeno"
        MyVM.data(26).name = "Yodo"
    Else
        nC = 0
        nf = FreeFile()
        Open App.Path & "\vitaname.txt" For Input As #nf
        Do While Not EOF(nf) And (nC < 27)
            Line Input #nf, sC
            MyVM.data(nC).name = Mid(Trim(sC), 1, 50)
            nC = nC + 1
        Loop
        Close #nf
    End If
    
    MyVM.data(0).type = "V"
    MyVM.data(1).type = "V"
    MyVM.data(2).type = "V"
    MyVM.data(3).type = "V"
    MyVM.data(4).type = "V"
    MyVM.data(5).type = "V"
    MyVM.data(6).type = "V"
    MyVM.data(7).type = "V"
    MyVM.data(8).type = "V"
    MyVM.data(9).type = "V"
    MyVM.data(10).type = "V"
    MyVM.data(11).type = "V"
    MyVM.data(12).type = "V"
    MyVM.data(13).type = "M"
    MyVM.data(14).type = "M"
    MyVM.data(15).type = "M"
    MyVM.data(16).type = "M"
    MyVM.data(17).type = "M"
    MyVM.data(18).type = "M"
    MyVM.data(19).type = "M"
    MyVM.data(20).type = "M"
    MyVM.data(21).type = "M"
    MyVM.data(22).type = "M"
    MyVM.data(23).type = "M"
    MyVM.data(24).type = "M"
    MyVM.data(25).type = "M"
    MyVM.data(26).type = "M"
    
    MyVM.data(0).units = "µg"
    MyVM.data(1).units = "µg"
    MyVM.data(2).units = "mg"
    MyVM.data(3).units = "µg"
    MyVM.data(4).units = "mg"
    MyVM.data(5).units = "mg"
    MyVM.data(6).units = "mg"
    MyVM.data(7).units = "mg"
    MyVM.data(8).units = "mg"
    MyVM.data(9).units = "µg"
    MyVM.data(10).units = "µg"
    MyVM.data(11).units = "µg"
    MyVM.data(12).units = "mg"
    MyVM.data(13).units = "mg"
    MyVM.data(14).units = "mg"
    MyVM.data(15).units = "mg"
    MyVM.data(16).units = "mg"
    MyVM.data(17).units = "mg"
    MyVM.data(18).units = "mg"
    MyVM.data(19).units = "mg"
    MyVM.data(20).units = "mg"
    MyVM.data(21).units = "mg"
    MyVM.data(22).units = "mg"
    MyVM.data(23).units = "µg"
    MyVM.data(24).units = "µg"
    MyVM.data(25).units = "µg"
    MyVM.data(26).units = "µg"
    
    MyVM.data(0).daily_ref = 800
    MyVM.data(1).daily_ref = 5
    MyVM.data(2).daily_ref = 12
    MyVM.data(3).daily_ref = 75
    MyVM.data(4).daily_ref = 80
    MyVM.data(5).daily_ref = 1.1
    MyVM.data(6).daily_ref = 1.4
    MyVM.data(7).daily_ref = 16
    MyVM.data(8).daily_ref = 1.4
    MyVM.data(9).daily_ref = 200
    MyVM.data(10).daily_ref = 2.5
    MyVM.data(11).daily_ref = 50
    MyVM.data(12).daily_ref = 6
    MyVM.data(13).daily_ref = 2000
    MyVM.data(14).daily_ref = 800
    MyVM.data(15).daily_ref = 800
    MyVM.data(16).daily_ref = 700
    MyVM.data(17).daily_ref = 375
    MyVM.data(18).daily_ref = 14
    MyVM.data(19).daily_ref = 10
    MyVM.data(20).daily_ref = 1
    MyVM.data(21).daily_ref = 2
    MyVM.data(22).daily_ref = 3.5
    MyVM.data(23).daily_ref = 55
    MyVM.data(24).daily_ref = 40
    MyVM.data(25).daily_ref = 50
    MyVM.data(26).daily_ref = 150

End Sub

''''''''''
'Nutricion
''''''''''
Public Function Capa1A_Item_Nutrition_Card_TK(Item As Item_Nutrition_Card) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.items where code=" & Item.Code, MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Item_Nutrition_Card_TK
    End If
    On Error GoTo FIN_Item_Nutrition_Card_TK
    If MyRecAux.EOF Then
        Resp = 1004
    End If
    If Resp <> 1004 Then
        MyRecAux.Fields("nutrition") = Item.Card
        MyRecAux.Update
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
FIN_Item_Nutrition_Card_TK:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Item_Nutrition_Card_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Item_Nutrition_Card_TK = Resp

End Function
Public Function Capa1A_Nutrition_Vita_Mine_TK(Vita_Mine As Nutrition_Vita_Mine, ByVal lBorrado As Boolean) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    
    On Error Resume Next
    If lBorrado = True Then
        MyConnObj.Execute "delete from pcscale.vitamins where card=" & CStr(Vita_Mine.Card) & " and identifier=" & CStr(Vita_Mine.identifier)
    Else
        MyRecAux.Open "select * from pcscale.nutrition where card=" & CStr(Vita_Mine.Card), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            GoTo FIN_Capa1A_Nutrition_Vita_Mine
        End If
        On Error GoTo FIN_Capa1A_Nutrition_Vita_Mine
        If MyRecAux.EOF Then
            Resp = 1004
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
        If Resp <> 1004 Then
            MyRecAux.Open "select * from pcscale.vitamins where card=" & CStr(Vita_Mine.Card) & " and identifier=" & CStr(Vita_Mine.identifier), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("card") = Vita_Mine.Card
                MyRecAux.Fields("identifier") = Vita_Mine.identifier
            End If
            MyRecAux.Fields("per100") = Vita_Mine.per100
            MyRecAux.Fields("perportion") = Vita_Mine.perportion
            MyRecAux.Fields("type") = MyVM.data(Vita_Mine.identifier).type
            MyRecAux.Fields("name") = MyVM.data(Vita_Mine.identifier).name
            MyRecAux.Fields("units") = MyVM.data(Vita_Mine.identifier).units
            MyRecAux.Fields("daily_ref") = MyVM.data(Vita_Mine.identifier).daily_ref
            MyRecAux.Update
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
    End If
FIN_Capa1A_Nutrition_Vita_Mine:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Nutrition_Vita_Mine_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Nutrition_Vita_Mine_TK = Resp

End Function

Public Function Capa1A_Nutrition_Base_TK(Card As Nutrition_Base, ByVal lBorrado As Boolean) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    
    On Error Resume Next
    If lBorrado Then
        MyConnObj.Execute "delete from pcscale.nutrition where card=" & CStr(Card.Card)
    Else
        MyRecAux.Open "select * from pcscale.nutrition where card=" & CStr(Card.Card), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            GoTo FIN_Capa1A_Nutrition
        End If
        On Error GoTo FIN_Capa1A_Nutrition
        If MyRecAux.EOF Then
            MyRecAux.AddNew
            MyRecAux.Fields("card") = Card.Card
        End If
        MyRecAux.Fields("carbohydrate_100") = Card.carbohydrate_100
        MyRecAux.Fields("carbohydrate_portion") = Card.carbohydrate_portion
        MyRecAux.Fields("energy_kcal_100") = Card.energy_kcal_100
        MyRecAux.Fields("energy_kcal_portion") = Card.energy_kcal_portion
        MyRecAux.Fields("energy_kj_100") = Card.energy_kj_100
        MyRecAux.Fields("energy_kj_portion") = Card.energy_kj_portion
        MyRecAux.Fields("fat_100") = Card.fat_100
        MyRecAux.Fields("fat_portion") = Card.fat_portion
        MyRecAux.Fields("fibre_100") = Card.fibre_100
        MyRecAux.Fields("fibre_portion") = Card.fibre_portion
        MyRecAux.Fields("mono_unsaturates_100") = Card.mono_unsaturates_100
        MyRecAux.Fields("mono_unsaturates_portion") = Card.mono_unsaturates_portion
        MyRecAux.Fields("polyols_100") = Card.polyols_100
        MyRecAux.Fields("polyols_portion") = Card.polyols_portion
        MyRecAux.Fields("polyunsaturates_100") = Card.polyunsaturates_100
        MyRecAux.Fields("polyunsaturates_portion") = Card.polyunsaturates_portion
        MyRecAux.Fields("portion_weight") = Card.portion_weight
        MyRecAux.Fields("protein_100") = Card.protein_100
        MyRecAux.Fields("protein_portion") = Card.protein_portion
        MyRecAux.Fields("salt_100") = Card.salt_100
        MyRecAux.Fields("salt_portion") = Card.salt_portion
        MyRecAux.Fields("saturates_100") = Card.saturates_100
        MyRecAux.Fields("saturates_portion") = Card.saturates_portion
        MyRecAux.Fields("starch_100") = Card.starch_100
        MyRecAux.Fields("starch_portion") = Card.starch_portion
        MyRecAux.Fields("sugars_100") = Card.sugars_100
        MyRecAux.Fields("sugars_portion") = Card.sugars_portion
        
        MyRecAux.Update
    End If
FIN_Capa1A_Nutrition:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Nutrition_Base_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Nutrition_Base_TK = Resp

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''

#Else


'Items
'-----
'Id   INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(127) NOT NULL
'ShortName Not VarChar(64)
'TexT Not VarChar(1024)
'Text1 Not VarChar(110)
'Text2 Not VarChar(110)
'Text3 Not VarChar(110)
'Text4 Not VarChar(110)
'Text5 Not VarChar(110)
'Family INTEGER
'Price DECIMAL(6,2) DEFAULT 0
'Discount Not DECIMAL(4, 2)
'SaleForm SMALLINT DEFAULT 1
'PLU INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'VAT SMALLINT DEFAULT 0
'EAN13 VarChar(13)
'Expiring SMALLINT DEFAULT 0
'BestBefore SMALLINT DEFAULT 0
'Tare DECIMAL(5,3) DEFAULT 0
'Label Not VarChar(64)
'Icon Not VarChar(127)
'QLevel Not VarChar(16)
'Unique Index(Code), Index(Name(19))'
' ShortName --> Descriptivo 1 == Name (Hasta 64 de longitud)
' Text1     --> Descriptivos 2 a 5
' Text2     --> Descriptivos 6 a 9
' Text3     --> Descriptivos 10 a 13
' Text4     --> Descriptivos 14 a 17
' Text5     --> Descriptivos 18 a 21
' Text      --> Se guarda en 2040 troceado de 80 en 80
' Discount  --> Se guarda en TIP_LETX
' Label     --> Se guarda en Tran_EP
' Icon y QLevel --> Se guardan en Imagen --> Icon &  " _+_ " & QLevel --> NO SE GUARDA QLevel
'

'PLUKeys
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Section INTEGER 'grupo en la Balanza
'Board INTEGER
'PLUKey INTEGER
'Code INTEGER
'Unique Index(Section, Board, PLUKey)'
'
'QLevels
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code VARCHAR(16) UNIQUE NOT NULL
'Icon Not VarChar(64)
'Description Not VarChar(127)''
'
'Vendors
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL'
'
'VendorKeys
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Counter INTEGER
'VendorKey INTEGER
'Vendor INTEGER UNIQUE
'Unique Index(Counter, VendorKey)'
'
'HTickets
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NetStat CHAR(1) DEFAULT ' '
'Captured CHAR(1) DEFAULT ' '
'tckType CHAR(1)
'status CHAR(1)
'Number INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'Scale INTEGER DEFAULT 0
'Vendor INTEGER
'DatTim DateTime
'ReloadFrom INTEGER DEFAULT 0
'Customer INTEGER DEFAULT 0
'tckDiscount DECIMAL(4,2) DEFAULT 0
'Unique Index(Vendor, id)'
'
'LTickets
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'IdHTicket INTEGER
'lineType SMALLINT
'SaleForm SMALLINT
'Item INTEGER
'Weight DECIMAL(5, 3)
'Tare DECIMAL(5,3) DEFAULT 0
'TTare CHAR(1) DEFAULT ' '
'lineDiscount DECIMAL(4, 2)
'Price DECIMAL(6, 2)
'Amount DECIMAL(8, 2)
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Index (IdHTicket)
'
'PayTickets
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'IdHTicket INTEGER
'PayMethod INTEGER
'Delivered DECIMAL(8,2) DEFAULT 0
'Index (IdHTicket)'
'
'PayMethods
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Name VARCHAR(64) UNIQUE NOT NULL'
'
'PacketLabels
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Number INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'Scale INTEGER DEFAULT 0
'Vendor INTEGER
'lineType SMALLINT
'Item INTEGER
'SaleForm SMALLINT
'Weight DECIMAL(5, 3)
'Tare DECIMAL(5,3) DEFAULT 0
'TTare CHAR(1) DEFAULT ' '
'Price DECIMAL(6, 2)
'lineDiscount DECIMAL(4, 2)
'Amount DECIMAL(8, 2)
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Index (IdHTicket)
'
'LocalVendorPacket
'-----------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Vendor INTEGER UNIQUE
'Auto BOOLEAN
'ClearStart BOOLEAN
'Repetitions SMALLINT DEFAULT 0
'Copies SMALLINT DEFAULT 0
'EAN13 CHAR(1) DEFAULT '1'
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Customer INTEGER DEFAULT 0
'PresetON BOOLEAN DEFAULT FALSE
'Item INTEGER DEFAULT 0
'SaleForm SMALLINT
'Price DECIMAL(6, 2)
'Units INTEGER UNSIGNED NOT NULL DEFAULT 0
'Tare DECIMAL(5, 3)
'Discount DECIMAL(4, 2)'
'
'LocalPacketPresets
'------------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Vendor INTEGER
'Level SMALLINT
'Type SMALLINT
'Value DECIMAL(6, 3)
'Number INTEGER DEFAULT 1
'ItemTotal INTEGER DEFAULT 0
'LowerLabels INTEGER UNSIGNED NOT NULL DEFAULT 0
'TotalUnits INTEGER UNSIGNED NOT NULL DEFAULT 0
'TotalWeight DECIMAL(12,3) UNSIGNED NOT NULL DEFAULT 0
'TotalAmount DECIMAL(12,2) UNSIGNED NOT NULL DEFAULT 0
'Unique Index(Vendor, Level)
'
'SServiceLabels
'--------------
'LIKE PacketLabels'
'
'NumTickets
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) DEFAULT 'T'
'Identifier INTEGER
'Number INTEGER
'UNIQUE INDEX (Identifier,Type)'
'
'EraseTickets
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) UNIQUE DEFAULT 'T'
'DatTim DateTime'
'
'NumberingTickets
'----------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) UNIQUE DEFAULT 'T'
'Numbering SMALLINT DEFAULT 0
'ResetNumber BOOLEAN DEFAULT FALSE
'SaveMode CHAR(1) DEFAULT ' ''
'
'Counters
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL
'Description Not VarChar(127)'
'
'Scales
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'MAC VARCHAR(17) UNIQUE NOT NULL
'IP Not VarChar(16)
'Number INTEGER
'Name VARCHAR(64) NOT NULL
'Type CHAR(1)'
'
'Families
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(54) NOT NULL
'IsBovine BOOLEAN DEFAULT FALSE
'Unique Index(Code)
'Label Not VarChar(64)
'Index (Name(19))'
' IsBovine  --> se guarda en Posicion
' Label     --> se guarda en Imagen

'
'Bovine
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NumCard SMALLINT(3) UNIQUE NOT NULL
'EarTag Not VarChar(30)
'BirthLand SMALLINT
'BirthDate Date
'BreedingLand SMALLINT
'SlaughterLand SMALLINT
'SlaughterReg VarChar(30)
'SlaughterDate Date
'ButcherLand SMALLINT
'ButcherReg VarChar(30)
'Lot VarChar(20)
'Index EarTag'
'
'Customers
'---------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(255) NOT NULL
'Address VarChar(255)
'City VarChar(255)
'WEB VarChar(255)
'Phone VarChar(255)
'TexT VarChar(255)
'Unique Index(Code)
'Index (Name(19))'
'
'Trademark
'---------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'CO_Name Not VarChar(80)
'Description Not VarChar(80)
'Business Not VarChar(80)
'Slogan Not VarChar(80)
'Headquarter Not VarChar(80)
'CO_WEB Not VarChar(80)
'Shop_Name Not VarChar(80)
'Address Not VarChar(80)
'City Not VarChar(80)
'Phone Not VarChar(80)
'TIN Not VarChar(80)
'Shop_WEB Not VarChar(80)
'Text1 Not VarChar(80)
'Text2 Not VarChar(80)
'Text3 Not VarChar(80)
'Text4 Not VarChar(80)
'Text5 Not VarChar(80)
'Text6 Not VarChar(80)'
'
'status
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Version SMALLINT DEFAULT 0'
'
'BarcodeEAN13
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NumberType CHAR(1)
'Number SMALLINT DEFAULT 0
'NetworkSales Not CHAR(12)
'NetworkLineSLB Not CHAR(12)
'NetworkTotalSLB Not CHAR(12)
'NetworkSS Not CHAR(12)
'NetworkPacket1 Not CHAR(12)
'NetworkPacket2 Not CHAR(12)
'LocalSales Not CHAR(12)
'LocalLineSLB Not CHAR(12)
'LocalTotalSLB Not CHAR(12)
'LocalSS Not CHAR(12)
'LocalPacket1 Not CHAR(12)
'LocalPacket2 Not CHAR(12)
'LocalTotal Not CHAR(12)
'Unique Index(NumberType, Number)'
'
'Users
'-----
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'User VARCHAR(16) UNIQUE NOT NULL DEFAULT ''
'password Not VarChar(31)
'Name VARCHAR(127) NOT NULL
'LevelLab BOOLEAN DEFAULT FALSE
'LevelAdmin BOOLEAN DEFAULT FALSE
'LevelConfig BOOLEAN DEFAULT FALSE
'LevelProg BOOLEAN DEFAULT FALSE
'LevelVendor BOOLEAN DEFAULT FALSE'
'
'-- Local tables
'LocalHTickets
'-------------
'LIKE HTickets'
'
'LocalLTickets
'-------------
'LIKE LTickets
'
'LocalPayTickets
'---------------
'LIKE PayTickets
'
'LocalPacketLabels
'-----------------
'LIKE PacketLabels
'
'LocalSServiceLabels
'-------------------
'LIKE SServiceLabels
'
'LocalNumTickets
'---------------
'LIKE NumTickets
'
'LocalEraseTickets
'-----------------
'LIKE EraseTickets
''
'
'LocalStatus
'-----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Version SMALLINT DEFAULT 0
'Language Not VarChar(31)
'ItemsUpdated BOOLEAN DEFAULT FALSE
'EraseTickets BOOLEAN DEFAULT FALSE
'RemoveScale BOOLEAN DEFAULT FALSE
'
'LocalCountries
'--------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(127) NOT NULL
'Unique Index(Code)
'Index (Name(19))

'Versión 4 de Touch-Scale no tiene Tabla Valores IVA's
'Public Sub Absorver_Iva()
'
'    Unload FrmIVA
'
'    CancelHumano = False
'    canCelar = 0
'    frmMonitor.cmdmonitorstop.Enabled = True
'
'    MostrarMonitor
'
'    frmControl.Switch_mode 0
'    If EstadoBal = 0 Then 'Protección mínima...
'        frmMonitor.MostrarInfo "Lectura Ivas..."
'        If canCelar = 0 And CancelHumano = False Then Call LeeIVAS
'    End If
'    'cerrar conexión
'    On Error Resume Next
'    MyConnObj.Close
'    Set MyConnObj = Nothing
'    MyRecSet.Close
'    Set MyRecSet = Nothing
'    On Error GoTo 0
'    EstadoBal = 1000
'    frmMonitor.cmdmonitorstop.Enabled = False
'    HabilitarSalir
'    frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
'    If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"'
'
'End Sub
'########################
' Nutricional
'########################

'*******************
'NUTRICION
'*******************
'DROP TABLE IF EXISTS `pcscale`.`nutrition`;
'CREATE TABLE  `pcscale`.`nutrition` (
'  `card` int(11) NOT NULL,
'  `portion_weight` float DEFAULT '0',
'  `energy_kj_100` float DEFAULT '0',
'  `energy_kj_portion` float DEFAULT '0',
'  `energy_kcal_100` float DEFAULT '0',
'  `energy_kcal_portion` float DEFAULT '0',
'  `fat_100` float DEFAULT '0',
'  `fat_portion` float DEFAULT '0',
'  `saturates_100` float DEFAULT '0',
'  `saturates_portion` float DEFAULT '0',
'  `mono_unsaturates_100` float DEFAULT '0',
'  `mono_unsaturates_portion` float DEFAULT '0',
'  `polyunsaturates_100` float DEFAULT '0',
'  `polyunsaturates_portion` float DEFAULT '0',
'  `carbohydrate_100` float DEFAULT '0',
'  `carbohydrate_portion` float DEFAULT '0',
'  `sugars_100` float DEFAULT '0',
'  `sugars_portion` float DEFAULT '0',
'  `polyols_100` float DEFAULT '0',
'  `polyols_portion` float DEFAULT '0',
'  `starch_100` float DEFAULT '0',
'  `starch_portion` float DEFAULT '0',
'  `fibre_100` float DEFAULT '0',
'  `fibre_portion` float DEFAULT '0',
'  `protein_100` float DEFAULT '0',
'  `protein_portion` float DEFAULT '0',
'  `salt_100` float DEFAULT '0',
'  `salt_portion` float DEFAULT '0',
'  PRIMARY KEY (`card`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
Public Type VitaMin
    daily_ref As Double
    name As String
    type As String
    units As String
End Type
Public Type D_VM
    data(26) As VitaMin
End Type
Public MyVM As D_VM

Public Type Item_Nutrition_Card
    Code As Long
    Card As Long
End Type

Public Type Nutrition_Base
    Card As Long
    portion_weight As Double
    energy_kj_100 As Double
    energy_kj_portion As Double
    energy_kcal_100 As Double
    energy_kcal_portion As Double
    fat_100 As Double
    fat_portion As Double
    saturates_100 As Double
    saturates_portion As Double
    mono_unsaturates_100 As Double
    mono_unsaturates_portion As Double
    polyunsaturates_100 As Double
    polyunsaturates_portion As Double
    carbohydrate_100 As Double
    carbohydrate_portion As Double
    sugars_100 As Double
    sugars_portion As Double
    polyols_100 As Double
    polyols_portion As Double
    starch_100 As Double
    starch_portion As Double
    fibre_100 As Double
    fibre_portion As Double
    protein_100 As Double
    protein_portion As Double
    salt_100 As Double
    salt_portion  As Double
End Type

Public Type Nutrition_Vita_Mine
    Card As Long
    identifier As Long '0 a 26 (abajo identificacion)
    'type As String 'V --> Vitamina, M --> Minerales
    'name As String 'abajo nombre en función identificacion
    'units As String 'abajo unidades en función identificacion
    'daily_ref As Double 'abajo referencia en funcion identificacion
    per100 As Double
    perportion As Double
End Type
'identifier type    name                units   daily_ref
'0          V       Vitamina A          ?g      800
'1          V       Vitamina D          ?g      5
'2          V       Vitamina E          mg      12
'3          V       Vitamina K          ?g      75
'4          V       Vitamina C          mg      80
'5          V       Tiamina             mg      1,1
'6          V       Riboflavina         mg      1,4
'7          V       Niacina             mg      16
'8          V       Vitamina B6         mg      1,4
'9          V       Ácido fólico        ?g      200
'10         V       Vitamina B12        ?g      2,5
'11         V       Biotina             ?g      50
'12         V       Ácido pantoténico   mg      6
'13         M       Potasio             mg      2000
'14         M       Cloro               mg      800
'15         M       Calcio              mg      800
'16         M       Fósforo             mg      700
'17         M       Magnesio            mg      375
'18         M       Hierro              mg      14
'19         M       Zinc                mg      10
'20         M       Cobre               mg      1
'21         M       Manganeso           mg      2
'22         M       Flúor               mg      3,5
'23         M       Selenio             ?g      55
'24         M       Cromo               ?g      40
'25         M       Molibdeno           ?g      50
'26         M       Yodo                ?g      150

Public Function DecodeUTF8(ByVal cnvUni As String)
If id = 6000 Then
    If cnvUni <> "" Then
        'Dim cnvUni2 As String
        'cnvUni2 = WToA(cnvUni)
        'CadenadeLog "cnvuni2:" & cnvUni2
        'DecodeUTF8 = AToW(cnvUni2)
        DecodeUTF8 = AToW(cnvUni)
    End If
Else
    DecodeUTF8 = cnvUni
End If
End Function

Public Function Item_Nutrition_Card_TK(Item As Item_Nutrition_Card) As Long
Dim Resp As Long
    CadenadeLog "Recibida orden de programación Ficha Nutricional a un artículo."
    Resp = 1010
    If Item.Code <> 0 And Item.Card >= 0 Then
        Resp = Capa1A_Item_Nutrition_Card_TK(Item)
    End If
    Item_Nutrition_Card_TK = Resp
End Function
'CCCCCCXXXXXX
Public Function ALT_Item_Nutrition_Card_TK(ByVal sItemCard As String) As Long
Dim Resp As Long
Dim Item As Item_Nutrition_Card

    CadenadeLog "Recibida orden ALT de programación Ficha Nutricional a un artículo."
    Resp = 1010
    
    Item.Code = left(sItemCard, 6)
    Item.Card = Mid(sItemCard, 7)
    
    If Item.Code <> 0 And Item.Card >= 0 Then
        Resp = Capa1A_Item_Nutrition_Card_TK(Item)
    End If
    ALT_Item_Nutrition_Card_TK = Resp
End Function

'**
Public Function Nutrition_Vita_Mine_TK(Vita_Mine As Nutrition_Vita_Mine)
Dim Resp As Long
    CadenadeLog "Recibida orden de programación Vitamina o mineral."
    Resp = 1010
    If Vita_Mine.Card > 0 Then
        Resp = Capa1A_Nutrition_Vita_Mine_TK(Vita_Mine, False)
    End If
    Nutrition_Vita_Mine_TK = Resp
End Function
'ccccccIIaaaaaabbbbbb
Public Function ALT_Nutrition_Vita_Mine_TK(ByVal sVita_Mine As String, ByVal lBorrado As Boolean)
Dim Resp As Long
Dim Vita_Mine As Nutrition_Vita_Mine
    CadenadeLog "Recibida orden ALT de programación Vitamina o mineral."
    Resp = 1010
    Vita_Mine.Card = left(sVita_Mine, 6)
    Vita_Mine.identifier = Mid(sVita_Mine, 7, 2)
    Vita_Mine.per100 = Val(Mid(sVita_Mine, 9, 6)) / 100
    Vita_Mine.perportion = Val(Mid(sVita_Mine, 15, 6)) / 100
    If Vita_Mine.Card > 0 Then
        Resp = Capa1A_Nutrition_Vita_Mine_TK(Vita_Mine, lBorrado)
    End If
    ALT_Nutrition_Vita_Mine_TK = Resp
End Function
Public Function Nutrition_Base_TK(Card As Nutrition_Base) As Long
Dim Resp As Long
    CadenadeLog "Recibida orden de programación Ficha Nutricional."
    Resp = 1010
    If Card.Card > 0 Then
        Resp = Capa1A_Nutrition_Base_TK(Card, False)
    End If
    Nutrition_Base_TK = Resp
End Function
'ALT_Nutrition_Vita_Mine_TK
'ALT_Nutrition_Base_TK
Public Function ALT_Nutrition_Base_TK(ByVal sCard As String, ByVal lBorrado As Boolean) As Long
Dim Card As Nutrition_Base
Dim Resp As Long

    CadenadeLog "Recibida orden ALT de programación Ficha Nutricional."
    Resp = 1010
    
    Card.Card = left(sCard, 6)
'  Correspondencia con Imagen
'  "portion_weight" -->1 (Peso por porción)
'  "energy_kj_100" -->2 (Valor energético en kj por 100g)
'  "energy_kj_portion" -->3 (Valor energético por porción en kj)
'  "energy_kcal_100" -->4 (Valor energético en kcal por 100g)
'  "energy_kcal_portion" -->5 (valor energético por porción en kcal)
'  "fat_100" -->6 (Total grasas en 100g)
'  "fat_portion" -->7 (Total grasas en porción)
'  "saturates_100" -->8 (Grasas saturadas en 100g)
'  "saturates_portion" -->9 (Grasas saturadas en porción)
'  "mono_unsaturates_100" -->10 (Grasas moinsaturadas en 100g)
'  "mono_unsaturates_portion" -->11 (Grasas monoinsaturadas en porción)
'  "polyunsaturates_100" -->12 (Grasas poliinsaturadas en 100g)
'  "polyunsaturates_portion" -->13 (Grasas poliinsaturadas en porción)
'  "carbohydrate_100" -->20 (Hidratos de carbono en 100g)
'  "carbohydrate_portion" -->21 (Hidratos de carbono en porción)
'  "sugars_100" -->22 (Azucares en 100g)
'  "sugars_portion" -->23 (Azucares en porción)
'  "polyols_100" -->24 (Polialcoholes en 100g)
'  "polyols_portion" -->25 (Polialcoholes en porción)
'  "starch_100" -->26 (Almidón en 100g)
'  "starch_portion" -->27 (Almidón en porción)
'  "fibre_100" -->14 (Fibra Alimentaria en 100g)
'  "fibre_portion" -->15 (Fibra Alimentaria en porción)
'  "protein_100" -->16 (Proteinas en 100g)
'  "protein_portion" -->17 (Proteinas en porción)
'  "salt_100" -->18 (Sal en 100g)
'  "salt_portion" -->19 (Sal en porción)
'Cada uno de los siguientes campos separados con "//" tienen 6 dígitos de longitud, siendo los 2 últimos dígitos los decimales
'carbohydrate_100//carbohydrate_portion//energy_kcal_100//energy_kcal_portion//energy_kj_100//energy_kj_portion//fat_100//
'fat_portion//fibre_100//fibre_portion//mono_unsaturates_100//mono_unsaturates_portion//polyols_100//polyols_portion//
'polyunsaturates_100//polyunsaturates_portion//portion_weight//protein_100//protein_portion//salt_100//salt_portion//
'saturates_100//saturates_portion//starch_100//starch_portion//sugars_100//sugars_portion
    Card.carbohydrate_100 = Val(Mid(sCard, 7, 6)) / 100
    Card.carbohydrate_portion = Val(Mid(sCard, 13, 6)) / 100
    Card.energy_kcal_100 = Val(Mid(sCard, 19, 6)) / 100
    Card.energy_kcal_portion = Val(Mid(sCard, 25, 6)) / 100
    Card.energy_kj_100 = Val(Mid(sCard, 31, 6)) / 100
    Card.energy_kj_portion = Val(Mid(sCard, 37, 6)) / 100
    Card.fat_100 = Val(Mid(sCard, 43, 6)) / 100
    Card.fat_portion = Val(Mid(sCard, 49, 6)) / 100
    Card.fibre_100 = Val(Mid(sCard, 55, 6)) / 100
    Card.fibre_portion = Val(Mid(sCard, 61, 6)) / 100
    Card.mono_unsaturates_100 = Val(Mid(sCard, 67, 6)) / 100
    Card.mono_unsaturates_portion = Val(Mid(sCard, 73, 6)) / 100
    Card.polyols_100 = Val(Mid(sCard, 79, 6)) / 100
    Card.polyols_portion = Val(Mid(sCard, 85, 6)) / 100
    Card.polyunsaturates_100 = Val(Mid(sCard, 91, 6)) / 100
    Card.polyunsaturates_portion = Val(Mid(sCard, 97, 6)) / 100
    Card.portion_weight = Val(Mid(sCard, 103, 6)) / 100
    Card.protein_100 = Val(Mid(sCard, 109, 6)) / 100
    Card.protein_portion = Val(Mid(sCard, 115, 6)) / 100
    Card.salt_100 = Val(Mid(sCard, 121, 6)) / 100
    Card.salt_portion = Val(Mid(sCard, 127, 6)) / 100
    Card.saturates_100 = Val(Mid(sCard, 133, 6)) / 100
    Card.saturates_portion = Val(Mid(sCard, 139, 6)) / 100
    Card.starch_100 = Val(Mid(sCard, 145, 6)) / 100
    Card.starch_portion = Val(Mid(sCard, 151, 6)) / 100
    Card.sugars_100 = Val(Mid(sCard, 157, 6)) / 100
    Card.sugars_portion = Val(Mid(sCard, 163, 6)) / 100
    
    If Card.Card > 0 Then
        Resp = Capa1A_Nutrition_Base_TK(Card, lBorrado)
    End If
    ALT_Nutrition_Base_TK = Resp

End Function

Public Sub carga_data_Vitamin()
Dim nf As Integer
Dim nC As Integer
Dim sC As String
    If Dir(App.Path & "\vitaname.txt") = "" Then
        MyVM.data(0).name = "Vitamina A"
        MyVM.data(1).name = "Vitamina D"
        MyVM.data(2).name = "Vitamina E"
        MyVM.data(3).name = "Vitamina K"
        MyVM.data(4).name = "Vitamina C"
        MyVM.data(5).name = "Tiamina"
        MyVM.data(6).name = "Riboflavina"
        MyVM.data(7).name = "Niacina"
        MyVM.data(8).name = "Vitamina B6"
        MyVM.data(9).name = "Ácido fólico"
        MyVM.data(10).name = "Vitamina B12"
        MyVM.data(11).name = "Biotina"
        MyVM.data(12).name = "Ácido pantoténico"
        MyVM.data(13).name = "Potasio"
        MyVM.data(14).name = "Cloro"
        MyVM.data(15).name = "Calcio"
        MyVM.data(16).name = "Fósforo"
        MyVM.data(17).name = "Magnesio"
        MyVM.data(18).name = "Hierro"
        MyVM.data(19).name = "Zinc"
        MyVM.data(20).name = "Cobre"
        MyVM.data(21).name = "Manganeso"
        MyVM.data(22).name = "Flúor"
        MyVM.data(23).name = "Selenio"
        MyVM.data(24).name = "Cromo"
        MyVM.data(25).name = "Molibdeno"
        MyVM.data(26).name = "Yodo"
    Else
        nC = 0
        nf = FreeFile()
        Open App.Path & "\vitaname.txt" For Input As #nf
        Do While Not EOF(nf) And (nC < 27)
            Line Input #nf, sC
            MyVM.data(nC).name = Mid(Trim(sC), 1, 50)
            nC = nC + 1
        Loop
        Close #nf
    End If
    
    MyVM.data(0).type = "V"
    MyVM.data(1).type = "V"
    MyVM.data(2).type = "V"
    MyVM.data(3).type = "V"
    MyVM.data(4).type = "V"
    MyVM.data(5).type = "V"
    MyVM.data(6).type = "V"
    MyVM.data(7).type = "V"
    MyVM.data(8).type = "V"
    MyVM.data(9).type = "V"
    MyVM.data(10).type = "V"
    MyVM.data(11).type = "V"
    MyVM.data(12).type = "V"
    MyVM.data(13).type = "M"
    MyVM.data(14).type = "M"
    MyVM.data(15).type = "M"
    MyVM.data(16).type = "M"
    MyVM.data(17).type = "M"
    MyVM.data(18).type = "M"
    MyVM.data(19).type = "M"
    MyVM.data(20).type = "M"
    MyVM.data(21).type = "M"
    MyVM.data(22).type = "M"
    MyVM.data(23).type = "M"
    MyVM.data(24).type = "M"
    MyVM.data(25).type = "M"
    MyVM.data(26).type = "M"
    
    MyVM.data(0).units = "µg"
    MyVM.data(1).units = "µg"
    MyVM.data(2).units = "mg"
    MyVM.data(3).units = "µg"
    MyVM.data(4).units = "mg"
    MyVM.data(5).units = "mg"
    MyVM.data(6).units = "mg"
    MyVM.data(7).units = "mg"
    MyVM.data(8).units = "mg"
    MyVM.data(9).units = "µg"
    MyVM.data(10).units = "µg"
    MyVM.data(11).units = "µg"
    MyVM.data(12).units = "mg"
    MyVM.data(13).units = "mg"
    MyVM.data(14).units = "mg"
    MyVM.data(15).units = "mg"
    MyVM.data(16).units = "mg"
    MyVM.data(17).units = "mg"
    MyVM.data(18).units = "mg"
    MyVM.data(19).units = "mg"
    MyVM.data(20).units = "mg"
    MyVM.data(21).units = "mg"
    MyVM.data(22).units = "mg"
    MyVM.data(23).units = "µg"
    MyVM.data(24).units = "µg"
    MyVM.data(25).units = "µg"
    MyVM.data(26).units = "µg"
    
    MyVM.data(0).daily_ref = 800
    MyVM.data(1).daily_ref = 5
    MyVM.data(2).daily_ref = 12
    MyVM.data(3).daily_ref = 75
    MyVM.data(4).daily_ref = 80
    MyVM.data(5).daily_ref = 1.1
    MyVM.data(6).daily_ref = 1.4
    MyVM.data(7).daily_ref = 16
    MyVM.data(8).daily_ref = 1.4
    MyVM.data(9).daily_ref = 200
    MyVM.data(10).daily_ref = 2.5
    MyVM.data(11).daily_ref = 50
    MyVM.data(12).daily_ref = 6
    MyVM.data(13).daily_ref = 2000
    MyVM.data(14).daily_ref = 800
    MyVM.data(15).daily_ref = 800
    MyVM.data(16).daily_ref = 700
    MyVM.data(17).daily_ref = 375
    MyVM.data(18).daily_ref = 14
    MyVM.data(19).daily_ref = 10
    MyVM.data(20).daily_ref = 1
    MyVM.data(21).daily_ref = 2
    MyVM.data(22).daily_ref = 3.5
    MyVM.data(23).daily_ref = 55
    MyVM.data(24).daily_ref = 40
    MyVM.data(25).daily_ref = 50
    MyVM.data(26).daily_ref = 150

End Sub

''''''''''
'Nutricion
''''''''''
Public Function Capa1A_Item_Nutrition_Card_TK(Item As Item_Nutrition_Card) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.items where code=" & Item.Code, MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Item_Nutrition_Card_TK
    End If
    On Error GoTo FIN_Item_Nutrition_Card_TK
    If MyRecAux.EOF Then
        Resp = 1004
    End If
    If Resp <> 1004 Then
        MyRecAux.Fields("nutrition") = Item.Card
        MyRecAux.Update
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
FIN_Item_Nutrition_Card_TK:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Item_Nutrition_Card_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Item_Nutrition_Card_TK = Resp

End Function
Public Function Capa1A_Nutrition_Vita_Mine_TK(Vita_Mine As Nutrition_Vita_Mine, ByVal lBorrado As Boolean) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    
    On Error Resume Next
    If lBorrado = True Then
        MyConnObj.Execute "delete from pcscale.vitamins where card=" & CStr(Vita_Mine.Card) & " and identifier=" & CStr(Vita_Mine.identifier)
    Else
        MyRecAux.Open "select * from pcscale.nutrition where card=" & CStr(Vita_Mine.Card), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            GoTo FIN_Capa1A_Nutrition_Vita_Mine
        End If
        On Error GoTo FIN_Capa1A_Nutrition_Vita_Mine
        If MyRecAux.EOF Then
            Resp = 1004
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
        If Resp <> 1004 Then
            MyRecAux.Open "select * from pcscale.vitamins where card=" & CStr(Vita_Mine.Card) & " and identifier=" & CStr(Vita_Mine.identifier), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("card") = Vita_Mine.Card
                MyRecAux.Fields("identifier") = Vita_Mine.identifier
            End If
            MyRecAux.Fields("per100") = Vita_Mine.per100
            MyRecAux.Fields("perportion") = Vita_Mine.perportion
            MyRecAux.Fields("type") = MyVM.data(Vita_Mine.identifier).type
            MyRecAux.Fields("name") = MyVM.data(Vita_Mine.identifier).name
            MyRecAux.Fields("units") = MyVM.data(Vita_Mine.identifier).units
            MyRecAux.Fields("daily_ref") = MyVM.data(Vita_Mine.identifier).daily_ref
            MyRecAux.Update
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
    End If
FIN_Capa1A_Nutrition_Vita_Mine:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Nutrition_Vita_Mine_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Nutrition_Vita_Mine_TK = Resp

End Function

Public Function Capa1A_Nutrition_Base_TK(Card As Nutrition_Base, ByVal lBorrado As Boolean) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    
    On Error Resume Next
    If lBorrado Then
        MyConnObj.Execute "delete from pcscale.nutrition where card=" & CStr(Card.Card)
    Else
        MyRecAux.Open "select * from pcscale.nutrition where card=" & CStr(Card.Card), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            GoTo FIN_Capa1A_Nutrition
        End If
        On Error GoTo FIN_Capa1A_Nutrition
        If MyRecAux.EOF Then
            MyRecAux.AddNew
            MyRecAux.Fields("card") = Card.Card
        End If
        MyRecAux.Fields("carbohydrate_100") = Card.carbohydrate_100
        MyRecAux.Fields("carbohydrate_portion") = Card.carbohydrate_portion
        MyRecAux.Fields("energy_kcal_100") = Card.energy_kcal_100
        MyRecAux.Fields("energy_kcal_portion") = Card.energy_kcal_portion
        MyRecAux.Fields("energy_kj_100") = Card.energy_kj_100
        MyRecAux.Fields("energy_kj_portion") = Card.energy_kj_portion
        MyRecAux.Fields("fat_100") = Card.fat_100
        MyRecAux.Fields("fat_portion") = Card.fat_portion
        MyRecAux.Fields("fibre_100") = Card.fibre_100
        MyRecAux.Fields("fibre_portion") = Card.fibre_portion
        MyRecAux.Fields("mono_unsaturates_100") = Card.mono_unsaturates_100
        MyRecAux.Fields("mono_unsaturates_portion") = Card.mono_unsaturates_portion
        MyRecAux.Fields("polyols_100") = Card.polyols_100
        MyRecAux.Fields("polyols_portion") = Card.polyols_portion
        MyRecAux.Fields("polyunsaturates_100") = Card.polyunsaturates_100
        MyRecAux.Fields("polyunsaturates_portion") = Card.polyunsaturates_portion
        MyRecAux.Fields("portion_weight") = Card.portion_weight
        MyRecAux.Fields("protein_100") = Card.protein_100
        MyRecAux.Fields("protein_portion") = Card.protein_portion
        MyRecAux.Fields("salt_100") = Card.salt_100
        MyRecAux.Fields("salt_portion") = Card.salt_portion
        MyRecAux.Fields("saturates_100") = Card.saturates_100
        MyRecAux.Fields("saturates_portion") = Card.saturates_portion
        MyRecAux.Fields("starch_100") = Card.starch_100
        MyRecAux.Fields("starch_portion") = Card.starch_portion
        MyRecAux.Fields("sugars_100") = Card.sugars_100
        MyRecAux.Fields("sugars_portion") = Card.sugars_portion
        
        MyRecAux.Update
    End If
FIN_Capa1A_Nutrition:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Nutrition_Base_TK Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_Nutrition_Base_TK = Resp

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''

'########################

Public Sub Absorver_Estructura()
                
    'Borrar datos previos existentes...
    'Revisar actual Hydra para ver netID
    CancelHumano = False
    canCelar = 0
    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = True
    frmMonitor.Label1.Caption = ""
    
    MostrarMonitor
    
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then 'Protección mínima...
        'frmMonitor.MostrarInfo CargaCadena(329) & " ..."  '"Secciones..."
        'Call LeeSecciones
        frmMonitor.MostrarInfo CargaCadena(303) & " ..."  '"Equipos..."
        If canCelar = 0 And CancelHumano = False Then Call LeeEquipos
        frmMonitor.MostrarInfo CargaCadena(67) & " ..." '"Familias..."
        If canCelar = 0 And CancelHumano = False Then Call LeeFamilias
        frmMonitor.MostrarInfo CargaCadena(13) & " ..." '"Vendedores..."
        If canCelar = 0 And CancelHumano = False Then Call LeeVendedores
        frmMonitor.MostrarInfo CargaCadena(214) & " ..." '"Ivas..."
        If canCelar = 0 And CancelHumano = False Then Call LeeIVAS
        frmMonitor.MostrarInfo CargaCadena(1084) & " ..." '"Articulos..."
        If canCelar = 0 And CancelHumano = False Then Call LeeArticulos
        'TouchScale ... prueba captura paises...
        frmMonitor.MostrarInfo CargaCadena(619) & " ..." '"Fichas de vacuno..."
        If canCelar = 0 And CancelHumano = False Then Call capturaPaises
        If canCelar = 0 And CancelHumano = False Then Call capturaFichas
        frmMonitor.MostrarInfo CargaCadena(23) & " ..." '"Códigos de barras..."
        If canCelar = 0 And CancelHumano = False Then Call capturaCB
        frmMonitor.MostrarInfo CargaCadena(1056) & " ..." '"Clientes..."
        If canCelar = 0 And CancelHumano = False Then Call capturaClientes
        ''''''''''''''''''''''''''''''''''''''''
        
        Call Corregir_Nulos
        'frmMonitor.MostrarInfo "Lectura Datos Trazabilidad..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTraza '(*)
        'frmMonitor.MostrarInfo "Lectura de Tipos de Etiquetas..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTipoEtq '(*)
        'frmMonitor.MostrarInfo "Lectura de Totales..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTickets
        
    End If
    'cerrar conexión
    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0
    EstadoBal = 1000
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.Command1.Enabled = True
    HabilitarSalir
    frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
    If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
End Sub

Public Function CodiIdent_FS(ByVal FS As Long) As Long
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim nFs As Long
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    Set myRec = Mybase.OpenRecordset("select * from seccion where secc_maqui=" & CStr(FS))
    If Not myRec.EOF Then
        nFs = myRec.Fields("codi_ident")
    Else
        myRec.Close
        Set myRec = Mybase.OpenRecordset("select * from seccion")
        If Not myRec.EOF Then
            myRec.MoveFirst
            nFs = myRec.Fields("codi_ident")
        End If
    End If
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    If Err.Number <> 0 Then
        CodiIdent_FS = 0
    Else
        CodiIdent_FS = nFs
    End If
    On Error GoTo 0
End Function

'Items
'-----
'DROP TABLE IF EXISTS `pcscale`.`items`;
'CREATE TABLE  `pcscale`.`items` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(127) NOT NULL,
'  `ShortName` varchar(64) NOT NULL,
'  `Text` varchar(1024) NOT NULL,
'  `Text1` varchar(110) NOT NULL,
'  `Text2` varchar(110) NOT NULL,
'  `Text3` varchar(110) NOT NULL,
'  `Text4` varchar(110) NOT NULL,
'  `Text5` varchar(110) NOT NULL,
'  `Family` int(11) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT '0.00',
'  `Discount` decimal(4,2) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `PLU` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `VAT` decimal(1,0) unsigned DEFAULT '0',
'  `EAN13` varchar(13) DEFAULT NULL,
'  `Expiring` smallint(6) DEFAULT '0',
'  `BestBefore` smallint(6) DEFAULT '0',
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `Label` varchar(64) NOT NULL,
'  `Icon` varchar(127) NOT NULL,
'  `QLevel` 8varchar(16) NOT NULL,
'  `Mixed` smallint(6) DEFAULT '0',
'  `MixedWeight` decimal(5,3) DEFAULT '0.000',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
'Repartir los descriptivos... Para posible compatibilidad...
' ShortName --> Descriptivo 1 == Name (Hasta 64 de longitud)
' name      -->
' Text1     -->
' Text2     -->
' Text3     -->
' Text4     -->
' Text5     --> Descriptivos 1 al 20
' Text      --> Se guarda en 2040 troceado de 80 en 80
Public Sub LeeArticulos()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset 'tabla articulo
    'Dim MyRecIva As DAO.Recordset
    Dim Myrec2040 As dao.Recordset
    Dim MyRecTmp As dao.Recordset
    Dim MyRecAux As New ADODB.Recordset
    
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim nPrc As Double
    'Dim n As Long
    Dim sDesc As String
    Dim nLen As Long
    Dim ncont As Double
    Dim sParte As String
    Dim a As Variant
    Dim nContEvent As Integer
    Dim lPrint As Boolean
    
    On Error Resume Next
    
    If Dir(App.Path & "\lprint") <> "" Then
        lPrint = True
    Else
        lPrint = False
    End If
    
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.items where code<>0 order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                
                nFam = MyRecSet.Fields("family")
                If nFam = 0 Then nFam = 1
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(nFam))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("codi_fam") = 1
                    myRec.Fields("codi_ident") = 0
                    myRec.Fields("secc_maqui") = 0
                    myRec.Fields("posicion") = 0
                    myRec.Fields("txt_fam") = "DPT. " & CStr(nFam)
                    myRec.Fields("imagen") = ""
                    myRec.Fields("etiqueta") = ""
                    myRec.Update
                End If
                myRec.Close
                Set myRec = Nothing
                
                Set myRec = Mybase.OpenRecordset("select * from articulo where codigo=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                
                
                nSec = MyRecSet.Fields("counter")
                nPlu = MyRecSet.Fields("plu")
                myRec.Fields("codi_ident") = CodiIdent_FS(nSec)
                If nSec = 0 Then
                    myRec.Fields("secc_maqui") = CodiIdent_FS(nSec)
                Else
                    myRec.Fields("secc_maqui") = nSec
                End If
                myRec.Fields("codi_sub") = 0
                myRec.Fields("codi_fam") = nFam
                myRec.Fields("plu") = nPlu
                myRec.Fields("codigo") = MyRecSet.Fields("code")
                If MyRecSet.Fields("saleform") = 1 Then
                    myRec.Fields("codi_pes") = "W"
                Else
                    myRec.Fields("codi_pes") = "U"
                End If
                
                
                myRec.Fields("tipo_iva") = MyRecSet.Fields("vat")
                
                nPrc = MyRecSet.Fields("price")
                If UsaEuro Then
                    myRec.Fields("euros") = nPrc
                Else
                    myRec.Fields("precio") = nPrc
                End If
                myRec.Fields("prc3") = Format(Now, "yymmddhhmm")
                myRec.Fields("tara") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("tara_envasado") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("caducidad") = MyRecSet.Fields("Expiring")
                myRec.Fields("pref") = MyRecSet.Fields("BestBefore")
                'Atención! Touch Scale no dispone de:
                '-Peso Aproximado
                '-Merma
                '-Undades de Preselección
                'MyRec.Fields("poid") = 0 'MyRecSet.Fields("pesoaproximado")
                myRec.Fields("fcb") = 0 'MyRecSet.Fields("merma")
                myRec.Fields("ning") = 0 'MyRecSet.Fields("unidadespreseleccion")
                'Atención! La etiqueta es por nombre... Se ha de crear un fichero
                ' en el PC con los nombres de etiquetas que existen en la Balanza
                ' para que se pueda asociar mediante un índice... O crear un campo
                ' en Articulo para contener el nombre...
                ' DE MOMENTO SE RECOGE NOMBRE DE ETIQUETA EN TRAN_EP (max 60 caracteres)
                myRec.Fields("etq") = 0 'ETQ se hace 0 hasta solucionar
                If IsNull(MyRecSet.Fields("label")) Then
                    myRec.Fields("tran_ep") = ""
                    myRec.Fields("etq") = 0
                Else
                    myRec.Fields("tran_ep") = Trim(DecodeUTF8(MyRecSet.Fields("label")))
                    Select Case DecodeUTF8(MyRecSet.Fields("label"))
                    Case "1.xml"
                        myRec.Fields("etq") = 1
                    Case "2.xml"
                        myRec.Fields("etq") = 2
                    Case "3.xml"
                        myRec.Fields("etq") = 3
                    Case "4.xml"
                        myRec.Fields("etq") = 4
                    Case "5.xml"
                        myRec.Fields("etq") = 5
                    Case "6.xml"
                        myRec.Fields("etq") = 6
                    Case "7.xml"
                        myRec.Fields("etq") = 7
                    Case "8.xml"
                        myRec.Fields("etq") = 8
                    Case "9.xml"
                        myRec.Fields("etq") = 9
                    Case "-"
                        myRec.Fields("etq") = 16
                    End Select
                    
                End If
                '1.0.2 ShortName no se captura...
                myRec.Fields("des_plu1") = Mid(Trim(DecodeUTF8(MyRecSet.Fields("name"))), 1, 40) 'era shortname
                '''''''''''''''''''''''''''''''''
                'Atención! no se dispone de Imagen ni Color
                If lPrint Then
                    myRec.Fields("art_cb") = Trim(DecodeUTF8(MyRecSet.Fields("ean13print")))
                Else
                    myRec.Fields("art_cb") = Trim(DecodeUTF8(MyRecSet.Fields("ean13")))
                End If
                myRec.Fields("grupo_cons") = 0 'Grupo de conservación
                myRec.Fields("imagen") = Trim(DecodeUTF8(MyRecSet.Fields("icon"))) '& " _+_ " & _
                                         'Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("qlevel"))))
                myRec.Fields("ning") = MyRecSet.Fields("discount")
                
                If MyRecSet.Fields("mixed") <> 0 Then
                    myRec.Fields("lmix") = True
                Else
                    myRec.Fields("lmix") = False
                End If
                myRec.Fields("poid") = 1000 * (MyRecSet.Fields("mixedweight"))
                If MyRecSet.Fields("onkey") <> 0 Then
                    myRec.Fields("onkey") = True
                Else
                    myRec.Fields("onkey") = False
                End If
                If Dir(App.Path & "\upper.txt") = "" Then
'''''''Ingredientes a tabla 2040
'                If Not IsNull(Trim(MyRecSet.Fields("text"))) Then
'                    Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("code")))
'                    If myRec.EOF Then
'                        Myrec2040.AddNew
'                    Else
'                        Myrec2040.Edit
'                    End If
'                    Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
'                    Myrec2040.Fields("plu") = nPlu
'                    Myrec2040.Fields("mostrador") = nSec
'                    Myrec2040.Fields("tran_txt2") = "-"
'                    sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text"))))
'                    nLen = Len(sDesc)
'                    ncont = nLen / 80
'
'                    If ncont <> Int(nLen / 80) Then ncont = Int(ncont) + 1
'
'                    For nLen = 1 To ncont
'
'                        Select Case nLen
'                        Case 1
'                            Myrec2040.Fields("txt_21") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 2
'                            Myrec2040.Fields("txt_22") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 3
'                            Myrec2040.Fields("txt_23") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 4
'                            Myrec2040.Fields("txt_24") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 5
'                            Myrec2040.Fields("txt_25") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 6
'                            Myrec2040.Fields("txt_26") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 7
'                            Myrec2040.Fields("txt_27") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 8
'                            Myrec2040.Fields("txt_28") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 9
'                            Myrec2040.Fields("txt_29") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 10
'                            Myrec2040.Fields("txt_30") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 11
'                            Myrec2040.Fields("txt_31") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 12
'                            Myrec2040.Fields("txt_32") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 13
'                            Myrec2040.Fields("txt_33") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 14
'                            Myrec2040.Fields("txt_34") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 15
'                            Myrec2040.Fields("txt_35") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 16
'                            Myrec2040.Fields("txt_36") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 17
'                            Myrec2040.Fields("txt_37") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 18
'                            Myrec2040.Fields("txt_38") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 19
'                            Myrec2040.Fields("txt_39") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 20
'                            Myrec2040.Fields("txt_40") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        End Select
'                    Next nLen
'                    Myrec2040.Update
'                    Myrec2040.Close
'                    Set Myrec2040 = Nothing
'                End If

'                'sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text1")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text2")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text3")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text4")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text5"))))
                
''''''Ingredientes a tabla 2040
                If Not IsNull(Trim(MyRecSet.Fields("text"))) Then
                    Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("code")))
                    If Myrec2040.EOF Then
                        Myrec2040.AddNew
                    Else
                        Myrec2040.Edit
                    End If
                    Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
                    Myrec2040.Fields("plu") = nPlu
                    Myrec2040.Fields("mostrador") = nSec
                    Myrec2040.Fields("tran_txt2") = "-"
                    sDesc = Trim(DecodeUTF8(MyRecSet.Fields("text")))
                    nLen = Len(sDesc)
                    ncont = nLen / 50
                    
                    If ncont <> Int(nLen / 50) Then ncont = Int(ncont) + 1
                        
                    For nLen = 1 To ncont
                        
                        Select Case nLen
                        Case 1
                            Myrec2040.Fields("txt_21") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 2
                            Myrec2040.Fields("txt_22") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 3
                            Myrec2040.Fields("txt_23") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 4
                            Myrec2040.Fields("txt_24") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 5
                            Myrec2040.Fields("txt_25") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 6
                            Myrec2040.Fields("txt_26") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 7
                            Myrec2040.Fields("txt_27") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 8
                            Myrec2040.Fields("txt_28") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 9
                            Myrec2040.Fields("txt_29") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 10
                            Myrec2040.Fields("txt_30") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 11
                            Myrec2040.Fields("txt_31") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 12
                            Myrec2040.Fields("txt_32") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 13
                            Myrec2040.Fields("txt_33") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 14
                            Myrec2040.Fields("txt_34") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 15
                            Myrec2040.Fields("txt_35") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 16
                            Myrec2040.Fields("txt_36") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 17
                            Myrec2040.Fields("txt_37") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 18
                            Myrec2040.Fields("txt_38") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 19
                            Myrec2040.Fields("txt_39") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        Case 20
                            Myrec2040.Fields("txt_40") = Mid(sDesc, ((nLen - 1) * 50) + 1, 50)
                        End Select
                    Next nLen
                    Myrec2040.Update
                    Myrec2040.Close
                    Set Myrec2040 = Nothing
                End If

                'sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))) & _
                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text1")))) & _
                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text2")))) & _
                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text3")))) & _
                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text4")))) & _
                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text5"))))
                sDesc = Trim(DecodeUTF8(MyRecSet.Fields("text1"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text2"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text3"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text4"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text5")))
                
                nLen = Len(sDesc)
                ncont = nLen / 25
                
                If ncont <> Int(nLen / 25) Then ncont = Int(ncont) + 1
                    
                For nLen = 1 To ncont
                    Select Case nLen
                    'Select Case MyRecAux.Fields("numlinea")
                        Case 1
                            myRec.Fields("des_plu2") = Trim(Mid(sDesc, 1, 25))
                        Case 2
                            myRec.Fields("des_plu3") = Trim(Mid(sDesc, 36, 25))
                        Case 3
                            myRec.Fields("des_plu4") = Trim(Mid(sDesc, 61, 25))
                        Case 4
                            myRec.Fields("des_plu5") = Trim(Mid(sDesc, 86, 25))
                        Case 5
                            myRec.Fields("des_plu6") = Trim(Mid(sDesc, 111, 25))
                        Case 6
                            myRec.Fields("des_plu7") = Trim(Mid(sDesc, 136, 25))
                        Case 7
                            myRec.Fields("des_plu8") = Trim(Mid(sDesc, 161, 25))
                        Case 8
                            myRec.Fields("des_plu9") = Trim(Mid(sDesc, 186, 25))
                        Case 9
                            myRec.Fields("des_plu0") = Trim(Mid(sDesc, 211, 25))
                        Case 10
                            myRec.Fields("des_plux") = Trim(Mid(sDesc, 236, 25))
                        Case 11
                            myRec.Fields("des_plu11") = Trim(Mid(sDesc, 261, 25))
                        Case 12
                            myRec.Fields("des_plu12") = Trim(Mid(sDesc, 286, 25))
                        Case 13
                            myRec.Fields("des_plu13") = Trim(Mid(sDesc, 311, 25))
                        Case 14
                            myRec.Fields("des_plu14") = Trim(Mid(sDesc, 336, 25))
                        Case 15
                            myRec.Fields("des_plu15") = Trim(Mid(sDesc, 361, 25))
                        Case 16
                            myRec.Fields("des_plu16") = Trim(Mid(sDesc, 386, 25))
                        Case 17
                            myRec.Fields("des_plu17") = Trim(Mid(sDesc, 411, 25))
                        Case 18
                            myRec.Fields("des_plu18") = Trim(Mid(sDesc, 436, 25))
                        Case 19
                            myRec.Fields("des_plu19") = Trim(Mid(sDesc, 461, 25))
                        Case 20
                            myRec.Fields("des_plu20") = Trim(Mid(sDesc, 486, 25))
                       
            
                    End Select
            
                Next nLen
                
                Else
                
                sDesc = Trim(DecodeUTF8(MyRecSet.Fields("text1"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text2"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text3"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text4"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text5"))) & _
                        Trim(DecodeUTF8(MyRecSet.Fields("text")))
                

                nLen = Len(sDesc)
                ncont = nLen / 25
                
                If ncont <> Int(nLen / 25) Then ncont = Int(ncont) + 1
                    
                For nLen = 1 To ncont
                    Select Case nLen
                    'Select Case MyRecAux.Fields("numlinea")
                        Case 1
                            myRec.Fields("des_plu2") = Trim(Mid(sDesc, 1, 25))
                        Case 2
                            myRec.Fields("des_plu3") = Trim(Mid(sDesc, 26, 25))
                        Case 3
                            myRec.Fields("des_plu4") = Trim(Mid(sDesc, 51, 25))
                        Case 4
                            myRec.Fields("des_plu5") = Trim(Mid(sDesc, 76, 25))
                        Case 5
                            myRec.Fields("des_plu6") = Trim(Mid(sDesc, 101, 25))
                        Case 6
                            myRec.Fields("des_plu7") = Trim(Mid(sDesc, 126, 25))
                        Case 7
                            myRec.Fields("des_plu8") = Trim(Mid(sDesc, 151, 25))
                        Case 8
                            myRec.Fields("des_plu9") = Trim(Mid(sDesc, 176, 25))
                        Case 9
                            myRec.Fields("des_plu0") = Trim(Mid(sDesc, 201, 25))
                        Case 10
                            myRec.Fields("des_plux") = Trim(Mid(sDesc, 226, 25))
                        Case 11
                            myRec.Fields("des_plu11") = Trim(Mid(sDesc, 251, 25))
                        Case 12
                            myRec.Fields("des_plu12") = Trim(Mid(sDesc, 276, 25))
                        Case 13
                            myRec.Fields("des_plu13") = Trim(Mid(sDesc, 301, 25))
                        Case 14
                            myRec.Fields("des_plu14") = Trim(Mid(sDesc, 326, 25))
                        Case 15
                            myRec.Fields("des_plu15") = Trim(Mid(sDesc, 351, 25))
                        Case 16
                            myRec.Fields("des_plu16") = Trim(Mid(sDesc, 376, 25))
                        Case 17
                            myRec.Fields("des_plu17") = Trim(Mid(sDesc, 401, 25))
                        Case 18
                            myRec.Fields("des_plu18") = Trim(Mid(sDesc, 426, 25))
                        Case 19
                            myRec.Fields("des_plu19") = Trim(Mid(sDesc, 451, 25))
                        Case 20
                            myRec.Fields("des_plu20") = Trim(Mid(sDesc, 476, 25))
                        Case Else
                            If nLen > 20 And nLen < 41 Then
                                Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("code")))
                                If myRec.EOF Then
                                    Myrec2040.AddNew
                                Else
                                    Myrec2040.Edit
                                End If
                                Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
                                Myrec2040.Fields("plu") = nPlu
                                Myrec2040.Fields("mostrador") = nSec
                                Myrec2040.Fields("tran_txt2") = "-"
                                Myrec2040.Fields("txt_" & CStr(nLen)) = Trim(Mid(sDesc, 476 + ((nLen - 20) * 25), 25))
                                Myrec2040.Update
                                Myrec2040.Close
                                Set Myrec2040 = Nothing
                            End If
                    End Select
            
                Next nLen
                    
                End If
                
                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                'Call frmControl.Balanza_ComOK
                nContEvent = nContEvent + 1
                If nContEvent >= 10 Then
                    nContEvent = 0
                    Do_Events
                End If
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
End Sub

'Families
'--------
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(54) NOT NULL,
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;

Public Sub LeeFamilias()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.families order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast

            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                'atención! V4 Touch Scale las Familias son Genéricas nos se asignan a sección.
                myRec.Fields("codi_fam") = MyRecSet.Fields("code")
                myRec.Fields("codi_ident") = 0 'MyRecSet.Fields("idseccion")
                myRec.Fields("secc_maqui") = 0 'MyRecSet.Fields("idseccion")
                If MyRecSet.Fields("isbovine") <> 0 Then
                    myRec.Fields("posicion") = 1
                    If MyRecSet.Fields("isbovine") = 0 Then
                        myRec.Fields("posicion") = 0
                    Else
                        myRec.Fields("posicion") = 1
                    End If
                Else
                    myRec.Fields("posicion") = 0 'MyRecSet.Fields("posicion")
                End If
                myRec.Fields("txt_fam") = Mid(Trim(DecodeUTF8(MyRecSet.Fields("name"))), 1, 35)
                myRec.Fields("etiqueta") = Trim(DecodeUTF8(MyRecSet.Fields("label")))
                myRec.Fields("imagen") = Trim(DecodeUTF8(MyRecSet.Fields("icon")))
                'atención! V4 Touch Scale en pcscale.families se almacena si es familia Vacuno y etiqueta
                '          a asignar a los artículos de la Familia. Se deberá recorrer la tabla de artículos
                '          para que los que coincidan con la Familia asignarles la Etiqueta. Por otra parte
                '          en Hydra se indica una sola familia de Vacuno en Configuración y se asigna número
                '          de etiqueta especial, habrá que asignar nombre en lugar de número.
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
        
        Else 'se crea la familia 1
        
            Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=1")
            If myRec.EOF Then
                myRec.AddNew
                myRec.Fields("codi_fam") = 1
                myRec.Fields("codi_ident") = 0
                myRec.Fields("secc_maqui") = 0
                myRec.Fields("posicion") = 0
                myRec.Fields("txt_fam") = "DPT. 1"
                myRec.Fields("imagen") = ""
                myRec.Fields("etiqueta") = ""
                myRec.Update
            End If
    
        End If
    End If
    ''''''''''''''''''''''''''''''''
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    ''''''''''''''''''''''''''''''''
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.counterfamilies order by countercode,familyindex", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast

            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from familias where secc_maqui=" & CStr(MyRecSet.Fields("countercode")) & " and codi_fam=" & CStr(MyRecSet.Fields("familycode")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                'atención! V4 Touch Scale las Familias son Genéricas nos se asignan a sección.
                myRec.Fields("codi_fam") = MyRecSet.Fields("familycode")
                myRec.Fields("codi_ident") = MyRecSet.Fields("countercode")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("countercode")
                myRec.Fields("tran_fam") = ""
                myRec.Fields("borrado") = False
                myRec.Fields("index") = MyRecSet.Fields("familyindex")
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
        
        End If
    End If
'''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
End Sub

'Vendors
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL'
'
'VendorKeys
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Counter INTEGER
'VendorKey INTEGER
'Vendor INTEGER UNIQUE
'Unique Index(Counter, VendorKey)'
'atención! V4 Touch Scale un vendedor no puede estar en varias Secciones
Public Sub LeeVendedores()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.vendorkeys order by vendor", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("vendor")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                
                myRec.Fields("ident_vend") = MyRecSet.Fields("vendor")
                myRec.Fields("codi_ident") = MyRecSet.Fields("counter")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("counter")
                myRec.Fields("tec_vend") = MyRecSet.Fields("vendorkey")
                myRec.Fields("tran_vend") = "*"
                
                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
            MyRecSet.Close
            Set MyRecSet = Nothing
            myRec.Close
            Set myRec = Nothing
            
            MyRecSet.Open "Select * from pcscale.vendors order by code", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                On Error GoTo 0
                EstadoBal = 1000
                canCelar = 1
                Call frmControl.Balanza_ComError(0)
            Else
'*******
                If Not MyRecSet.EOF Then
                    MyRecSet.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = MyRecSet.Recordcount
                    
                    MyRecSet.MoveFirst
                    'n = 1
                    Do While Not MyRecSet.EOF
                        Set myRec = Mybase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("code")))
                        
                        If Not myRec.EOF Then
                            myRec.MoveFirst
                            Do While Not myRec.EOF
                                myRec.Edit
                    
                                myRec.Fields("codi_vend") = MyRecSet.Fields("code")
                                myRec.Fields("nombre") = Mid(DecodeUTF8(MyRecSet.Fields("name")), 1, 30)
                                myRec.Fields("prog_vend") = "*"
                                 
                                myRec.Update
                                myRec.Movenext
                            Loop
                             
                        End If
                        
                        MyRecSet.Movenext
                        
                        Call MonitorBarra
                        Call frmControl.Balanza_ComOK
                        
                        Do_Events
                        
                        If canCelar = 1 Or CancelHumano = True Then GoTo salir
                    
                    Loop
                End If
'*******
            End If
            
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
    
    
End Sub

'Counters
'--------
'DROP TABLE IF EXISTS `pcscale`.`counters`;
'CREATE TABLE  `pcscale`.`counters` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Description` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
Public Sub LeeSecciones()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim cm As String
    
    
    On Error Resume Next
    
    If lTransPlu Then Exit Sub
    
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select * from pcscale.counters where code<>0 order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast


            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
        
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                If MyMinSec = 0 Then
                    MyMinSec = MyRecSet.Fields("code")
                End If
                Set myRec = Mybase.OpenRecordset("select * from seccion where codi_ident=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                myRec.Fields("com") = 0
                myRec.Fields("sec_ip") = ""
                myRec.Fields("codi_ident") = MyRecSet.Fields("code")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("code")
                myRec.Fields("nombre") = Mid(DecodeUTF8(MyRecSet.Fields("name")), 1, 25)
                myRec.Fields("descripcio") = Mid(DecodeUTF8(MyRecSet.Fields("description")), 1, 25)
                'atención! V4 touch scale se dejará com GA y se verá si cambiar...
                myRec.Fields("enviardatos") = "GA"
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
End Sub

'Scales
'------
'DROP TABLE IF EXISTS `pcscale`.`scales`;
'CREATE TABLE  `pcscale`.`scales` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `MAC` varchar(17) NOT NULL,
'  `IP` varchar(16) NOT NULL,
'  `Number` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Type` char(1) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `MAC` (`MAC`)
') ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
'Tabla scales:
'.- Type: ‘S’, ‘N’, ‘A’
'    Balanza servidora (‘S’), esclava (‘N’), balanza no “conectada” (‘A’)
Public Sub LeeEquipos()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim MyRecSec As dao.Recordset
    Dim sData As String
    Dim cm As String
    
    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select * from pcscale.scales order by number", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
   
   
        If Not MyRecSet.EOF Then
            frmMonitor.barra1.Value = 0
            MyRecSet.MoveLast

            
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from equipos where numero_eqp=" & CStr(MyRecSet.Fields("number")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                myRec.Fields("numero_eqp") = MyRecSet.Fields("number")
                sData = Trim(DecodeUTF8(MyRecSet.Fields("ip")))
                If GeneraTextoIP(sData) Then
                    myRec.Fields("prog_eqp") = sData
                End If
                'If MyRecSet.Fields("tipoconexion") = 1 Then
                '    MyRec.Fields("modelo") = 0
                'Else
                    myRec.Fields("modelo") = MyRecSet.Fields("type")
                'End If
                myRec.Fields("descripcio") = Mid(Trim(DecodeUTF8(MyRecSet.Fields("name"))), 1, 21)
                If MyRecSet.Fields("counter") <> 0 Then
                    myRec.Fields("codi_ident") = MyRecSet.Fields("counter")
                    myRec.Fields("secc_maqui") = MyRecSet.Fields("counter")
                Else
                    If MyMinSec <> 0 Then
                        myRec.Fields("codi_ident") = MyMinSec
                        myRec.Fields("secc_maqui") = MyMinSec
                    Else
                        myRec.Fields("codi_ident") = 1
                        myRec.Fields("secc_maqui") = 1
                    End If
                End If
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
            
   
        End If
        
        
    End If
salir:
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
fin:
    On Error GoTo 0
End Sub




'DROP TABLE IF EXISTS `pcscale`.`vats`;
'CREATE TABLE  `pcscale`.`vats` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` decimal(1,0) unsigned NOT NULL,
'  `Name` varchar(16) NOT NULL,
'  `Percent` decimal(4,2) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
Public Sub LeeIVAS()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset 'tabla articulo
    Dim MyRecAux As New ADODB.Recordset
    'Dim nFam As Long
    'Dim nSec As Long
    'Dim nPlu As Long
    'Dim n As Long

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    'n = 0
    MyRecAux.CursorLocation = adUseClient
    MyRecAux.Open "Select * from pcscale.vats order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecAux.EOF Then
            MyRecAux.MoveLast
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecAux.Recordcount
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
             'MyRecSet.Open "Select * from dbo.tipoimpuesto where idtipoimpuesto=" & CStr(MyRecAux.Fields("idtipoimpuesto1")), MyConnObj, adOpenKeyset, adLockOptimistic
             'If Not MyRecSet.EOF Then
                Set myRec = Mybase.OpenRecordset("select * from iva where tipo=" & CStr(MyRecAux.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("tipo") = MyRecAux.Fields("code")
                   'n = n + 1
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("code")
                   myRec.Fields("descripcion") = DecodeUTF8(MyRecAux.Fields("name"))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                   'If n > 4 Then
                   '   GoTo Salir
                   'End If
                Else
                '    n = n + 1
                   myRec.Edit
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("code")
                   myRec.Fields("descripcion") = DecodeUTF8(MyRecAux.Fields("name"))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                
                End If
             'End If
             MyRecAux.Movenext
             'MyRecSet.Close
             'Set MyRecSet = Nothing
             Call MonitorBarra
             Call frmControl.Balanza_ComOK

             Do_Events

             If canCelar = 1 Or CancelHumano = True Then GoTo salir

            Loop

        End If

    End If
    ''''''''''''''''''''''''''''''''

salir:
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    Mybase.Close
    Set Mybase = Nothing
fin:
    On Error GoTo 0

End Sub

'DROP TABLE IF EXISTS `pcscale`.`localcountries`;
'CREATE TABLE  `pcscale`.`localcountries` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Alpha2` char(2) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Alpha2` (`Alpha2`),
'  UNIQUE KEY `Code` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=245 DEFAULT CHARSET=utf8;
Public Sub capturaPaises()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    On Error GoTo SalErrPaises
    MyRecSet.Open "Select * from pcscale.localcountries order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Mybase.Execute ("delete from paises")
        Set MyRecVar = Mybase.OpenRecordset("select * from paises")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            If Not IsNull(MyRecSet.Fields("alpha2")) Then
                If Len(Trim(MyRecSet.Fields("alpha2"))) >= 2 Then
                    'On Error GoTo Siguientepais
                    MyRecVar.AddNew
                    MyRecVar.Fields("codigo") = MyRecSet.Fields("code")
                    MyRecVar.Fields("nombre") = Mid(DecodeUTF8(MyRecSet.Fields("name")), 1, 100) 'Format(MyRecSet.Fields("code"), "000") & "-" & Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("alpha2"))), 1, 2) & "-" & "   -" & Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 89)
                    MyRecVar.Fields("valor") = Mid(DecodeUTF8(MyRecSet.Fields("name")), 1, 100)
                    MyRecVar.Fields("tran_pais") = " "
                    MyRecVar.Fields("borrado") = False
                    MyRecVar.Fields("modificado") = True
                    MyRecVar.Fields("abrv_3") = "   "
                    'On Error Resume Next
                    MyRecVar.Fields("abrv_2") = Mid(DecodeUTF8(MyRecSet.Fields("alpha2")), 1, 2)
                    'On Error GoTo 0
                    MyRecVar.Fields("tipo") = 1
                    MyRecVar.Update
'Siguientepais:
                    'If Err.Number <> 0 Then
                    '    CadenadeLog "error en paises: " & MyRecVar.Fields("abrv_2") & "ERR:" & CStr(Err.Number) & " -->" & Err.Description
                    '    On Error GoTo 0
                    '
                    'End If
                End If
            End If
            MyRecSet.Movenext
        Loop
SalErrPaises:
        On Error GoTo 0
        MyRecVar.Close
        Set MyRecVar = Nothing
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'DROP TABLE IF EXISTS `pcscale`.`bovine`;
'CREATE TABLE  `pcscale`.`bovine` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumCard` smallint(3) NOT NULL,
'  `EarTag` varchar(30) NOT NULL, --> Crotal
'  `BirthLand` smallint(6) DEFAULT NULL, --> Pais Nacimiento
'  `BirthDate` date DEFAULT NULL, --> Fecha nacimiento
'  `BreedingLand` smallint(6) DEFAULT NULL, --> Pais Crianza
'  `SlaughterLand` smallint(6) DEFAULT NULL,--> Pais Sacrificio
'  `SlaughterReg` varchar(30) DEFAULT NULL, --> Registro Sacrificio
'  `SlaughterDate` date DEFAULT NULL, --> Fecha Sacrificio
'  `ButcherLand` smallint(6) DEFAULT NULL, --> Pais Despiece
'  `ButcherReg` varchar(30) DEFAULT NULL, --> Registro despiece
'  `Lot` varchar(20) DEFAULT NULL, --> Lote
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumCard` (`NumCard`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub capturaFichas()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    MyRecSet.Open "Select * from pcscale.bovine order by numcard", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno where codigo=" & Chr(34) & CStr(MyRecSet.Fields("numcard")) & Chr(34))
            If MyRecVar.EOF Then
                MyRecVar.AddNew
                'MyRecVar.Fields("codigo") = MyRecSet.Fields("numcard")
                'MyRecVar.Fields("identificador") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("eartag")))
                'MyRecVar.Fields("codnacimiento") = MyRecSet.Fields("birthland")
                'MyRecVar.Fields("fecha_nac") = Format(MyRecSet.Fields("birthdate"), "dd/mm/yyyy")
                'MyRecVar.Fields("codcrianza") = MyRecSet.Fields("breedingland")
                'MyRecVar.Fields("codsacrificio") = MyRecSet.Fields("slaughterland")
                'MyRecVar.Fields("rsisacrificio") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("SlaughterReg")))
                'MyRecVar.Fields("fecha_sac") = Format(MyRecSet.Fields("SlaughterDate"), "dd/mm/yyyy")
                'MyRecVar.Fields("coddespiece") = MyRecSet.Fields("ButcherLand")
                'MyRecVar.Fields("rsidespiece") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("ButcherReg")))
                'MyRecVar.Fields("lote") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("lot")))
            Else
                MyRecVar.Edit
            End If
                MyRecVar.Fields("codigo") = MyRecSet.Fields("numcard")
                MyRecVar.Fields("identificador") = DecodeUTF8(MyRecSet.Fields("eartag"))
                MyRecVar.Fields("codnacimiento") = MyRecSet.Fields("birthland")
                MyRecVar.Fields("fecha_nac") = Format(MyRecSet.Fields("birthdate"), "dd/mm/yyyy")
                MyRecVar.Fields("codcrianza") = MyRecSet.Fields("breedingland")
                MyRecVar.Fields("codsacrificio") = MyRecSet.Fields("slaughterland")
                MyRecVar.Fields("rsisacrificio") = DecodeUTF8(MyRecSet.Fields("SlaughterReg"))
                MyRecVar.Fields("fecha_sac") = Format(MyRecSet.Fields("SlaughterDate"), "dd/mm/yyyy")
                MyRecVar.Fields("coddespiece") = MyRecSet.Fields("ButcherLand")
                MyRecVar.Fields("rsidespiece") = DecodeUTF8(MyRecSet.Fields("ButcherReg"))
                If IsNull(MyRecSet.Fields("lot")) Then
                    MyRecVar.Fields("lote") = ""
                Else
                    MyRecVar.Fields("lote") = DecodeUTF8(MyRecSet.Fields("lot"))
                End If
            'End If
            
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

Public Function consultaSeccMaquiBalanza(ByVal nBal) As Integer
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select secc_maqui from equipos where numero_eqp=" & CStr(nBal))
    If MyRecVar.EOF Then
        consultaSeccMaquiBalanza = 1
    Else
        consultaSeccMaquiBalanza = MyRecVar.Fields("secc_maqui")
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
End Function
'DROP TABLE IF EXISTS `pcscale`.`barcodeean13`;
'CREATE TABLE  `pcscale`.`barcodeean13` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumberType` char(1) DEFAULT NULL,
'  `Number` smallint(6) DEFAULT '0',
'  `NetworkSales` char(12) NOT NULL,
'  `NetworkLineSLB` char(12) NOT NULL,
'  `NetworkTotalSLB` char(12) NOT NULL,
'  `NetworkSS` char(12) NOT NULL,
'  `NetworkPacket1` char(12) NOT NULL,
'  `NetworkPacket2` char(12) NOT NULL,
'  `LocalSales` char(12) NOT NULL,
'  `LocalLineSLB` char(12) NOT NULL,
'  `LocalTotalSLB` char(12) NOT NULL,
'  `LocalSS` char(12) NOT NULL,
'  `LocalPacket1` char(12) NOT NULL,
'  `LocalPacket2` char(12) NOT NULL,
'  `LocalTotal` char(12) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumberType` (`NumberType`,`Number`)
') ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
'Typo "G" se desprecia, se cogen "S" (Equipo) y "C" (Mostrador)
Public Sub capturaCB()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim nCodiIdent As Integer
Dim nSeccMaqui As Integer
Dim nC As Integer
Dim cm As String

    Set Mybase = dao.OpenDatabase(Base_General)
    ''''''''''''''''''''''''''
    cm = "delete from codbar"
    Mybase.Execute cm
    ''''''''''''''''''''''''''
    MyRecSet.Open "Select * from pcscale.barcodeean13 where numbertype<>'G'", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from codbar")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            If MyRecSet.Fields("numbertype") = "S" Then
                nSeccMaqui = consultaSeccMaquiBalanza(MyRecSet.Fields("number"))
                nCodiIdent = nSeccMaqui
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where numero_eqp=" & CStr(MyRecSet.Fields("number")))
            Else
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where secc_maqui=" & CStr(MyRecSet.Fields("number")) & " and isnull(numero_eqp)")
                nSeccMaqui = MyRecSet.Fields("number")
                nCodiIdent = nSeccMaqui
            End If
            If MyRecVar.EOF Then
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.AddNew
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("numbertype") = "S" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = DecodeUTF8(MyRecSet.Fields("networksales"))
                        MyRecVar.Fields("CB_SUP") = DecodeUTF8(MyRecSet.Fields("NetworkLineSLB"))
                        MyRecVar.Fields("CB_MIX") = DecodeUTF8(MyRecSet.Fields("NetworkTotalSLB"))
                        MyRecVar.Fields("CB_EV1") = DecodeUTF8(MyRecSet.Fields("NetworkSS"))
                        MyRecVar.Fields("CB_EV2") = DecodeUTF8(MyRecSet.Fields("NetworkPacket1"))
                    Else
                        MyRecVar.Fields("CB_VNT") = DecodeUTF8(MyRecSet.Fields("localsales"))
                        MyRecVar.Fields("CB_SUP") = DecodeUTF8(MyRecSet.Fields("localLineSLB"))
                        MyRecVar.Fields("CB_MIX") = DecodeUTF8(MyRecSet.Fields("localTotalSLB"))
                        MyRecVar.Fields("CB_EV1") = DecodeUTF8(MyRecSet.Fields("localSS"))
                        MyRecVar.Fields("CB_EV2") = DecodeUTF8(MyRecSet.Fields("localPacket1"))
                    End If
                    MyRecVar.Update
                Next nC
            Else
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.Edit
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("numbertype") = "S" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = DecodeUTF8(MyRecSet.Fields("networksales"))
                        MyRecVar.Fields("CB_SUP") = DecodeUTF8(MyRecSet.Fields("NetworkLineSLB"))
                        MyRecVar.Fields("CB_MIX") = DecodeUTF8(MyRecSet.Fields("NetworkTotalSLB"))
                        MyRecVar.Fields("CB_EV1") = DecodeUTF8(MyRecSet.Fields("NetworkSS"))
                        MyRecVar.Fields("CB_EV2") = DecodeUTF8(MyRecSet.Fields("NetworkPacket1"))
                    Else
                        MyRecVar.Fields("CB_VNT") = DecodeUTF8(MyRecSet.Fields("localsales"))
                        MyRecVar.Fields("CB_SUP") = DecodeUTF8(MyRecSet.Fields("localLineSLB"))
                        MyRecVar.Fields("CB_MIX") = DecodeUTF8(MyRecSet.Fields("localTotalSLB"))
                        MyRecVar.Fields("CB_EV1") = DecodeUTF8(MyRecSet.Fields("localSS"))
                        MyRecVar.Fields("CB_EV2") = DecodeUTF8(MyRecSet.Fields("localPacket1"))
                    End If
                    MyRecVar.Update
                Next nC
    
            End If
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(255) NOT NULL,
'  `Address` varchar(255) DEFAULT NULL,
'  `City` varchar(255) DEFAULT NULL,
'  `WEB` varchar(255) DEFAULT NULL,
'  `Phone` varchar(255) DEFAULT NULL,
'  `Text` varchar(255) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub capturaClientes()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
    On Error GoTo sal_CapturaClientes
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
    'Mybase.Execute "delete from cliente"
    'If MyRecSet.State <> 0 Then
        'MyRecSet.Close
        'Set MyRecSet = Nothing
    'End If
    'MyRecSet.Open "Select * from pcscale.customers order by tin", MyConnObj, adOpenKeyset, adLockOptimistic
    MyRecSet.Open "Select * from pcscale.customers order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            'Set MyRecVar = Mybase.OpenRecordset("select * from cliente where cod_cli=" & CStr(MyRecSet.Fields("tin")))
            'Set MyRecVar = Mybase.OpenRecordset("select * from cliente where cod_cli=" & CStr(MyRecSet.Fields("id")))
            Set MyRecVar = Mybase.OpenRecordset("select * from cliente where ucase(trim(data3))='" & UCase(Trim(MyRecSet.Fields("tin"))) & "'")
            If MyRecVar.EOF Then
                MyRecVar.AddNew
                'MyRecVar.Fields("cod_cli") = MyRecSet.Fields("tin")
                MyRecVar.Fields("cod_cli") = MyRecSet.Fields("id")
                MyRecVar.Fields("nom_cli") = DecodeUTF8(MyRecSet.Fields("name"))
                MyRecVar.Fields("data1") = DecodeUTF8(MyRecSet.Fields("address"))
                MyRecVar.Fields("data2") = DecodeUTF8(MyRecSet.Fields("city"))
                'MyRecVar.Fields("data3") = DecodeUTF8(MyRecSet.Fields("web"))
                MyRecVar.Fields("data3") = DecodeUTF8(MyRecSet.Fields("tin"))
                MyRecVar.Fields("data4") = DecodeUTF8(MyRecSet.Fields("phone"))
            Else
                MyRecVar.Edit
                'MyRecVar.Fields("cod_cli") = MyRecSet.Fields("tin")
                MyRecVar.Fields("cod_cli") = MyRecSet.Fields("id")
                MyRecVar.Fields("nom_cli") = DecodeUTF8(MyRecSet.Fields("name"))
                MyRecVar.Fields("data1") = DecodeUTF8(MyRecSet.Fields("address"))
                MyRecVar.Fields("data2") = DecodeUTF8(MyRecSet.Fields("city"))
                'MyRecVar.Fields("data3") = DecodeUTF8(MyRecSet.Fields("web"))
                MyRecVar.Fields("data3") = DecodeUTF8(MyRecSet.Fields("tin"))
                MyRecVar.Fields("data4") = DecodeUTF8(MyRecSet.Fields("phone"))
            End If
            
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close
sal_CapturaClientes:
    If Err.Number <> 0 Then
        CadenadeLog "Error en CapturaClientes -->Err:" & CStr(Err.Number) & " desc.:" & Err.Description
    End If
    On Error GoTo 0
End Sub


Public Function verLF(ByVal sS As String) As String
Dim sR As String
Dim sF As String
Dim nL As Integer
Dim nC As Integer
Dim nP As Integer
    If lTransPlu Then
        'sR = Replace(sR, "CONS", Chr(10) & "CONS")
        'sR = Replace(sR, "Ingre", Chr(10) & "Ingre")
        'sR = Replace(sR, "ingre", Chr(10) & "Ingre")
        sR = Replace(sS, Chr(94), Chr(10))

        sR = Replace(sR, "  ", " ")
        sR = Replace(sR, "    ", " ")
        sR = Replace(sR, "      ", " ")
        
        sF = ""
        nL = Len(sR)
        nP = 0
        For nC = 1 To nL
            If Mid(sR, nC, 1) = "|" Then
                If nP = 0 Then
                    If nC < nL Then
                        sF = sF & "<b>"
                    End If
                    nP = 1
                Else
                    sF = sF & "</b>"
                    nP = 0
                End If
            Else
                sF = sF & Mid(sR, nC, 1)
            End If
        Next nC
        If nP = 1 Then
            sF = sF & "</b>"
        End If
        If sF <> "" Then
            sR = sF
        End If
        'verLF = sR
        
        'sR = Replace(sR, ",|", ", |")
        'sR = Replace(sR, " |", " <b>")
        ''If InStr(1, sR, "b>") <> 0 Then
        ''    sR = sR
        ''End If
        'verLF = Replace(sR, "|", "</b>")
        
    Else
        sR = Replace(sS, "/~", "</b>")
        sR = Replace(sR, "~", "<b>")
        sR = Replace(sR, Chr(94), Chr(10))
        'verLF =
        sR = Replace(sR, "|", Chr(10))
    End If
'Replace(Replace(Replace(Replace(Replace(Replace(Replace(myRecTmp1.Fields("text1"), "&", " "), Chr(13) & Chr(10), Chr(10)), "^", Chr(10)), "[", "<b><u>"), "]", "</u></b>"), "–", "-"), ";", ",")
    'verLF = Replace(Replace(Replace(Replace(Replace(Replace(sr, "&", " "), Chr(13) & Chr(10), Chr(10)), "^", Chr(10)), "[", "<b><u>"), "]", "</u></b>"), "–", "-"), ";", ",")
    'sR = Replace(sR, "&", " ")
    'sR = Replace(Replace(sR, "&", " "), Chr(13) & Chr(10), Chr(10))
    'sR = Replace(Replace(Replace(sR, "&", " "), Chr(13) & Chr(10), Chr(10)), "^", Chr(10))
    'sR = Replace(Replace(Replace(Replace(sR, "&", " "), Chr(13) & Chr(10), Chr(10)), "^", Chr(10)), "[", "<b><u>")
    verLF = Replace(Replace(Replace(Replace(Replace(sR, "&", " "), Chr(13) & Chr(10), Chr(10)), "^", Chr(10)), "[", "<b><u>"), "]", "</u></b>")
    
    
End Function
'//////////
'Envios Touch Scale
'\\\\\\\\\\

Public Function Item_OrderInFamily_T(ByVal nFam As Long, ByVal nMethod As Long) As Long
'nFam == número de familia
'nMethod --> 0=Código Artículo, 1=Plu Artículo, 2=Name, 3=shortname
'SELECT code,family,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=3 order by shortname
Dim Resp As Long
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
Dim n As Long
    Resp = 9999
    On Error GoTo Sal_T
    Select Case nMethod
    Case 0
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 1
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by plu", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 2
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by name", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 3
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by shortname", MyConnObj, adOpenKeyset, adLockOptimistic
    Case Else
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    End Select
    Resp = 10000
    n = 0
    If Not MyRecDes.EOF Then
        MyRecDes.MoveFirst
        Do While Not MyRecDes.EOF
            Resp = 10001
            cm = "update items set familyorder=" & CStr(MyRecDes.Fields("rank")) & " where code=" & CStr(MyRecDes.Fields("code")) & " and family=" & CStr(nFam)
            CadenadeLog cm
            MyConnObj.Execute cm
            n = n + 1
            MyRecDes.Movenext
        Loop
    End If
    Resp = 10002
    MyRecDes.Close
    Set MyRecDes = Nothing
Sal_T:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'Resp = 10000
    Else
        Resp = n
    End If
    Item_OrderInFamily_T = Resp
End Function

''''''''''''''''''''''
'ENVIO DE ARTÍCULOS...Touch
''''''''''''''''''''''
'Se envían:
'   IVAS
'   Familias Altas/bajas
'   Artículos Altas/bajas
'   Teclas (Siempre) Altas/bajas
'Atención PLU MIXTO
'DROP TABLE IF EXISTS `pcscale`.`items`;
'CREATE TABLE  `pcscale`.`items` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(127) NOT NULL,
'  `ShortName` varchar(64) NOT NULL,
'  `Text` varchar(1024) NOT NULL,
'  `Text1` varchar(110) NOT NULL,
'  `Text2` varchar(110) NOT NULL,
'  `Text3` varchar(110) NOT NULL,
'  `Text4` varchar(110) NOT NULL,
'  `Text5` varchar(110) NOT NULL,
'  `Family` int(11) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT '0.00',
'  `Discount` decimal(4,2) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `PLU` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `VAT` decimal(1,0) unsigned DEFAULT '0',
'  `EAN13` varchar(13) DEFAULT NULL,
'  `Expiring` smallint(6) DEFAULT '0',
'  `BestBefore` smallint(6) DEFAULT '0',
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `Label` varchar(64) NOT NULL,
'  `Icon` varchar(127) NOT NULL,
'  `QLevel` varchar(16) NOT NULL,
'  `Mixed` smallint(6) DEFAULT '0',
'  `MixedWeight` decimal(5,3) DEFAULT '0.000',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;

'//*
'DROP TABLE IF EXISTS `pcscale`.`counterfamilies`;
'CREATE TABLE  `pcscale`.`counterfamilies` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `CounterCode` int(11) NOT NULL,
'  `FamilyIndex` int(11) NOT NULL DEFAULT '1',
'  `FamilyCode` int(11) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `idcofam` (`CounterCode`,`FamilyIndex`),
'  CONSTRAINT `fkcofam` FOREIGN KEY (`CounterCode`) REFERENCES `counters` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB AUTO_INCREMENT=138 DEFAULT CHARSET=utf8;
'*//
Public Function MUE_EnviaArticuloNW(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean, Optional MiCodigo As Long) As Boolean
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
Dim myRecTmp1 As dao.Recordset
Dim myrecfam As dao.Recordset
Dim MyRecFam1 As dao.Recordset
Dim MyRecSec As dao.Recordset
Dim nD As Double
Dim MyRecAux As New ADODB.Recordset
Dim MyRecAux1 As New ADODB.Recordset
Dim MyRecAux2 As New ADODB.Recordset
Dim MyRecFamSec As New ADODB.Recordset
Dim cm As String
Dim nf As Integer
Dim nL As Integer
Dim lEsAlta As Boolean
Dim nC(99) As Long
Dim ncont As Integer
Dim Resp As Long
Dim nPrice As Double
Dim nContPrecios As Integer
Dim nIdFinal As Double
Dim sColorTmp As String
Dim PosiRel As Long
Dim lEnvImage As Boolean
Dim nContEvnt As Integer
Dim s(5) As String
Dim lEsTest As Boolean
Dim lNewTEXT As Boolean
Dim cm1 As String
Dim lEnvShort As Boolean
Dim l40TXT As Boolean
Dim l39TXT As Boolean
Dim l5l As Boolean
Dim l2140Text1 As Boolean
Dim nFI As Long
Dim nTaxFam(1000) As Integer
Dim sLabel1(1000) As String
Dim slabel2(1000) As String

Dim nL1 As Integer

'QUITAR raspberry
'On Error Resume Next
'raspberrY:
'cm = "delete from pcscale.items"
'MyConnObj.Execute cm
'On Error GoTo 0
''''''''''''''''''

Dim sXcb As String

l5l = False
l40TXT = False
l39TXT = False
l2140Text1 = False
If (Dir(App.Path & "\40l.txt") <> "") Or lTransPlu Then
    l40TXT = True
    If (Dir(App.Path & "\l2140text1") <> "") Then
        l2140Text1 = True
    End If

End If
If (Dir(App.Path & "\39l.txt") <> "") Then
    l39TXT = True
End If
If Dir(App.Path & "\5l.txt") <> "" Then
    l5l = True
End If

'CadenadeLog "Entrada en Envio de Articulos..."
'''''''''''''''''''
'Call Corregir_Nulos
'''''''''''''''''
If Dir(App.Path & "\test.tst") <> "" Then
    lEsTest = True
End If
lEnvShort = False
If Dir(App.Path & "\shname.txt") <> "" Then
    lEnvShort = True
End If

If l40TXT = False And l5l = False And l39TXT = False Then
    l40TXT = True
End If

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINARTICULO
    End If
    On Error GoTo 0
    'On Error Resume Next
    'Set myRecTmp1 = Mybase.OpenRecordset("select * from text15")
    'If Err.Number <> 0 Then
    '    lNewTEXT = False
    'Else
        lNewTEXT = True
    'End If
    'On Error GoTo 0
    'myRecTmp1.Close
    'Set myRecTmp1 = Nothing
    
'''''Nutricional
'ALT_Nutrition_Vita_Mine_TK
'ALT_Nutrition_Base_TK
    For nL1 = 0 To 99
        nTaxFam(nL1) = 0
    Next nL1
    For nL1 = 0 To 99
        sLabel1(nL1) = ""
    Next nL1
    For nL1 = 0 To 99
        slabel2(nL1) = ""
    Next nL1
    
    Set myrecfam = Mybase.OpenRecordset("select * from fam_code")
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
            If Not IsNull(myrecfam.Fields("tax")) Then
                nTaxFam(myrecfam.Fields("codi_fam")) = myrecfam.Fields("tax")
            End If
            If Not IsNull(myrecfam.Fields("label1")) And Trim(myrecfam.Fields("label1")) <> "" Then
                sLabel1(myrecfam.Fields("codi_fam")) = myrecfam.Fields("label1")
            End If
            If Not IsNull(myrecfam.Fields("label1")) And Trim(myrecfam.Fields("label2")) <> "" Then
                slabel2(myrecfam.Fields("codi_fam")) = myrecfam.Fields("label2")
            End If
            If Not myrecfam.EOF Then myrecfam.Movenext
        Loop
        
    End If
    myrecfam.Close
    Set myrecfam = Nothing

On Error GoTo SigueTrasNutri
If lAutoRecEnv = False Then
    If Modificaciones Then
        Set myRecTmp1 = Mybase.OpenRecordset("select * from nutrition where tran_nutri=' '")
    Else
        Set myRecTmp1 = Mybase.OpenRecordset("select * from nutrition")
    End If
    If Not myRecTmp1.EOF Then
        myRecTmp1.MoveFirst
        Do While Not myRecTmp1.EOF
            If myRecTmp1.Fields("Borrado") = "*" Then
               ALT_Nutrition_Base_TK myRecTmp1.Fields("cadena"), True
               Mybase.Execute "delete from vitamins where card=" & myRecTmp1.Fields("card")
            Else
                ALT_Nutrition_Base_TK myRecTmp1.Fields("cadena"), False
            End If
            myRecTmp1.Movenext
        Loop
        Mybase.Execute "delete from nutrition where borrado='*'"
    End If
    myRecTmp1.Close
    Set myRecTmp1 = Nothing
    If Modificaciones Then
        Set myRecTmp1 = Mybase.OpenRecordset("select * from vitamins where tran_vita=' '")
    Else
        Set myRecTmp1 = Mybase.OpenRecordset("select * from vitamins")
    End If
    If Not myRecTmp1.EOF Then
        myRecTmp1.MoveFirst
        Do While Not myRecTmp1.EOF
            If myRecTmp1.Fields("Borrado") = "*" Then
               ALT_Nutrition_Vita_Mine_TK myRecTmp1.Fields("cadena"), True
            Else
               ALT_Nutrition_Vita_Mine_TK myRecTmp1.Fields("cadena"), False
            End If
            myRecTmp1.Movenext
        Loop
        Mybase.Execute "delete from vitamins where borrado='*'"
    End If
    myRecTmp1.Close
    Set myRecTmp1 = Nothing

'''''''''''''''''''''''''''''''
SigueTrasNutri:
    On Error GoTo 0
    
    On Error GoTo FINARTICULO
''''''''''''''''''''''''
'    Call capturaClientes
'''''''PROGRAMACIÓN IVA's

        ''DROP TABLE IF EXISTS `pcscale`.`itempricerates`;
        ''CREATE TABLE  `pcscale`.`itempricerates` (
        ''  `Id` int(11) NOT NULL AUTO_INCREMENT,
        ''  `Code` int(11) NOT NULL,
        ''  `Rate` decimal(1,0) NOT NULL DEFAULT '1',
        ''  `Price` decimal(8,2) NOT NULL DEFAULT '0.00',
        ''  `Step` decimal(5,3) NOT NULL DEFAULT '0.000',
        ''  PRIMARY KEY (`Id`),
        ''  UNIQUE KEY `idx_rate` (`Code`,`Rate`,`Step`),
        ''  KEY `Rate` (`Rate`),
        ''  CONSTRAINT `itempricerates_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE,
        ''  CONSTRAINT `itempricerates_ibfk_2` FOREIGN KEY (`Rate`) REFERENCES `pricerates` (`Code`) ON DELETE CASCADE
        '') ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
        'CREATE TABLE `pricerates` (
        '  `Code` decimal(1,0) NOT NULL,
        '  `Name` varchar(64) NOT NULL DEFAULT '',
        '  PRIMARY KEY (`Code`)
        ') ENGINE=InnoDB DEFAULT CHARSET=utf8;
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=1", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 1
                MyRecAux.Fields("name") = "Tar. 1"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=2", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 2
                MyRecAux.Fields("name") = "Tar. 2"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing '
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=3", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 3
                MyRecAux.Fields("name") = "Tar. 3"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing '
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=4", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 4
                MyRecAux.Fields("name") = "Tar. 4"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing '
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=5", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 5
                MyRecAux.Fields("name") = "Tar. 5"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=6", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 6
                MyRecAux.Fields("name") = "Tar. 6"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=7", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 7
                MyRecAux.Fields("name") = "Tar. 7"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=8", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 8
                MyRecAux.Fields("name") = "Tar. 8"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.pricerates where code=9", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then GoTo SigueTec1
            On Error GoTo 0
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                MyRecAux.Fields("code") = 9
                MyRecAux.Fields("name") = "Tar. 9"
                MyRecAux.Update
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            'On Error Resume Next
            'MyRecAux.Open "select * from pcscale.pricerates where code=10", MyConnObj, adOpenKeyset, adLockOptimistic
            'If Err.Number <> 0 Then GoTo SigueTec1
            'On Error GoTo 0
            'On Error GoTo FINARTICULO
            'If MyRecAux.EOF Then
            '    MyRecAux.AddNew
            '    MyRecAux.Fields("code") = 10
            '    MyRecAux.Fields("name") = "Tar. 10"
            '    MyRecAux.Update
            'End If
            'MyRecAux.Close
            'Set MyRecAux = Nothing


SigueTec1:
'CadenadeLog "Verificación IVA..."

'ALT_Nutrition_Vita_Mine_TK
'ALT_Nutrition_Base_TK

CadenadeLog "Cuenca 1"
    For nL1 = 0 To 99
        If nTaxFam(nL1) <> 0 Then
            MyRecAux.Open "select * from pcscale.vats where code=" & CStr(nTaxFam(nL1)), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            
            MyRecAux.Fields("code") = nTaxFam(nL1)
            MyRecAux.Fields("name") = " VAT" & CStr(nTaxFam(nL1)) 'Trim(ConvertStringToUtf8Bytes(myRecTmp1.Fields("descripcion")))
            MyRecAux.Fields("percent") = 0 'myRecTmp1.Fields("valor")
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If
    Next nL1
CadenadeLog "Cuenca 2"
GoTo mio
    Set myRecTmp1 = Mybase.OpenRecordset("select * from iva")
    If Not myRecTmp1.EOF Then
        myRecTmp1.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = myRecTmp1.Recordcount
        myRecTmp1.MoveFirst
        Do While Not myRecTmp1.EOF
            MyRecAux.Open "select * from pcscale.vats where code=" & CStr(myRecTmp1.Fields("tipo")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            End If
            
            MyRecAux.Fields("code") = myRecTmp1.Fields("tipo")
            MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes(myRecTmp1.Fields("descripcion")))
            MyRecAux.Fields("percent") = myRecTmp1.Fields("valor")
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            myRecTmp1.Movenext
        
            Call MonitorBarra
        
            'Do_Events '--> ver
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 20 Then
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
        
        Loop
    
    End If
    myRecTmp1.Close
    Set myRecTmp1 = Nothing
'''''''''''''''''''''''''''''''
mio:

    'Programación Familias...
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(54) NOT NULL,
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
    'baja de familias...

'CadenadeLog "Baja Familias Fam_Code..."
'    For nL1 = 0 To 99
'        nTaxFam(nL1) = 0
'    Next nL1
'    Set myrecfam = Mybase.OpenRecordset("select * from fam_code")
'    If Not myrecfam.EOF Then
'
'        myrecfam.MoveFirst
'
'        Do While Not myrecfam.EOF
'
'            nTaxFam(myrecfam.Fields("codi_fam")) = myrecfam.Fields("tax")
'            If Not myrecfam.EOF Then myrecfam.Movenext
'        Loop
'
'    End If
'    myrecfam.Close
'    Set myrecfam = Nothing
CadenadeLog "Cuenca 3"


    Set myrecfam = Mybase.OpenRecordset("select * from fam_code where posicion=1525")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
        
            cm = "delete from pcscale.families where code=" & CStr(myrecfam.Fields("codi_fam"))
            MyConnObj.Execute cm
            cm = "delete from pcscale.items where family=" & CStr(myrecfam.Fields("codi_fam"))
            MyConnObj.Execute cm
            '1.5.7
            cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
            MyConnObj.Execute cm
            cm = "delete from familias where codi_fam=" & CStr(myrecfam.Fields("codi_fam"))
            Mybase.Execute cm
            ''''''
            
            myrecfam.Delete
            
            If Not myrecfam.EOF Then myrecfam.Movenext
        Loop
        
    End If
    myrecfam.Close
    Set myrecfam = Nothing
    ''''''''''''''''''''
    
'CadenadeLog "Alta familias Fam_Code..."

    Set myrecfam = Mybase.OpenRecordset("select * from fam_code order by codi_fam")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
    If Not myrecfam.EOF Then
        
        myrecfam.MoveFirst
        
        Do While Not myrecfam.EOF
            
            lEsAlta = False
            
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.families where code=" & CStr(myrecfam.Fields("codi_fam")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            If MyRecAux.EOF Then
                lEsAlta = True
            End If
            
            If lEsAlta Then
                'CadenadeLog "Se detecta familia:" & CStr(myrecfam.Fields("codi_fam")) & " No existente en Tabla Balanza..."
                MyRecAux.AddNew
            End If
            
            MyRecAux.Fields("code") = myrecfam.Fields("codi_fam")
            If IsNull(myrecfam.Fields("txt_fam")) Then
                MyRecAux.Fields("name") = "Familia " & CStr(myrecfam.Fields("codi_fam"))
            Else
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("txt_fam")))
            End If
            On Error Resume Next
            If myrecfam.Fields("etiqueta") <> "" Then
                'If Mid(Trim(ConvertStringToUtf8Bytes(MyRecFam.Fields("imagen"))), 1, 1) = "#" Then
                '    MyRecAux.Fields("label") = ""
                'Else
                '    If lEsTest Then
                '        If InStr(1, MyRecFam.Fields("imagen"), ".jpg") <> 0 Or InStr(1, MyRecFam.Fields("imagen"), ".png") <> 0 Then
                '            MyRecAux.Fields("label") = ""
                '            MyRecAux.Fields("icon") = Trim(ConvertStringToUtf8Bytes(MyRecFam.Fields("imagen")))
                '        Else
                            MyRecAux.Fields("label") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("etiqueta"))) 'era imagen
                '        End If
                '    Else
                '        MyRecAux.Fields("label") = Trim(ConvertStringToUtf8Bytes(MyRecFam.Fields("imagen")))
                '    End If
                'End If
            Else
                If lEsAlta Then MyRecAux.Fields("label") = ""
                'If IsNull(MyRecAux.Fields("label")) Or (MyRecAux.Fields("label") = "") Then
                '    MyRecAux.Fields("label") = ""
                'End If
            End If
            If myrecfam.Fields("imagen") <> "" Then
                MyRecAux.Fields("icon") = Trim(ConvertStringToUtf8Bytes(myrecfam.Fields("imagen"))) 'era imagen
            Else
                If lEsAlta Then MyRecAux.Fields("icon") = ""
                'If IsNull(MyRecAux.Fields("icon")) Then
                '    MyRecAux.Fields("icon") = ""
                'End If
            End If
            
            If myrecfam.Fields("posicion") = 1 Then
                MyRecAux.Fields("isbovine") = 1
            Else
                MyRecAux.Fields("isbovine") = 0
            End If
            'MyRecAux.UpdateBatch adAffectAll
            
            'CadenadeLog "Se graba..."
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'If lEsTest Then
            '    'asignación de indice
            '    If (myrecfam.Fields("secc_maqui")) <> 0 Then
            '        MyRecAux.Open "select * from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
            '        If MyRecAux.EOF Then
            '            MyRecAux.AddNew
            '            MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
            '            MyRecAux.Fields("familyindex") = myrecfam.Fields("codi_fam")
            '            MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
            '            MyRecAux.Update
            '        End If
            '        MyRecAux.Close
            '        Set MyRecAux = Nothing
            '    Else
            '        If (myrecfam.Fields("codi_fam")) <> 0 Then
            '            Set MyRecFam1 = Mybase.OpenRecordset("select codi_fam,secc_maqui from articulo where codi_fam=" & CStr(myrecfam.Fields("codi_fam")))
            '            If Not MyRecFam1.EOF Then
            '                MyRecFam1.MoveFirst
            '                MyRecAux.Open "select * from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
            '                If MyRecAux.EOF Then
            '                    MyRecAux.AddNew
            '                    MyRecAux.Fields("countercode") = MyRecFam1.Fields("secc_maqui")
            '                    MyRecAux.Fields("familyindex") = myrecfam.Fields("codi_fam")
            '                    MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
            '                    MyRecAux.Update
            '                End If
            '                MyRecAux.Close
            '                Set MyRecAux = Nothing
            '            End If
            '            MyRecFam1.Close
            '            Set MyRecFam1 = Nothing
            '        End If
            '    End If
            'End If
            
            myrecfam.Movenext
            Do_Events
        Loop
    
        
    End If
    'MyRecAux.Fields("descripcion") = MyRecFam.Fields("txt_fam")
    Do_Events
    myrecfam.Close
    Set myrecfam = Nothing
    
'CadenadeLog "Alta familias Fam_Code Finalizada..."

    'If lUpperNW Or Dir(App.Path & "\cambioposicion") <> "" Then
    If lUpper Or Dir(App.Path & "\cambioposicion") <> "" Then
        If (Not Modificaciones) Or (Dir(App.Path & "\cambioposicion") <> "") Then
            cm = "delete from pcscale.counterfamilies"
            MyConnObj.Execute cm
        End If
        If Dir(App.Path & "\cambioposicion") <> "" Then
            Kill App.Path & "\cambioposicion"
        End If
        
        Set myrecfam = Mybase.OpenRecordset("select * from familias")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
        If Not myrecfam.EOF Then
            myrecfam.MoveFirst
            Do While Not myrecfam.EOF
                If myrecfam.Fields("borrado") = True Then
                    '1.5.7
                    cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
                    MyConnObj.Execute cm
                    ''''''
                    myrecfam.Delete
                Else
                
                    If myrecfam.Fields("index") = 0 Then
                        cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
                        MyConnObj.Execute cm
                    End If
                    
                    'asignación de indice
                    MyRecAux.Open "select * from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                        MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
                        MyRecAux.Fields("familyindex") = 999999
                        MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
                        MyRecAux.Update
                    End If
                    MyRecAux.Close
                    Set MyRecAux = Nothing
        
                    Set MyRecFam1 = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(myrecfam.Fields("codi_fam")) & " and index<>0")
                    If Not MyRecFam1.EOF Then
                        'If lUpperNW Then
                        '    cm = "update pcscale.counterfamilies set familyindex=" & CStr(MyRecFam1.Fields("index") Mod (MyRecFam.Fields("secc_maqui") * 100)) & " where familycode=" & CStr(MyRecFam.Fields("codi_fam"))
                        'Else
                            cm = "update pcscale.counterfamilies set familyindex=" & CStr(MyRecFam1.Fields("index")) & " where familycode=" & CStr(myrecfam.Fields("codi_fam"))
                        'End If
                        MyConnObj.Execute cm
                    End If
        
                    MyRecFam1.Close
                    Set MyRecFam1 = Nothing
        
                End If
                If Not myrecfam.EOF Then myrecfam.Movenext
            Loop
        End If
        myrecfam.Close
        Set myrecfam = Nothing
        
    Else
    
'CadenadeLog "Alta familias 'Familias' a counter Families..."

        ''baja de familias en mostradores...
        'Set myrecfam = Mybase.OpenRecordset("select * from familias")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
        'If Not myrecfam.EOF Then
        '    myrecfam.MoveFirst
        '    Do While Not myrecfam.EOF
        '        If myrecfam.Fields("borrado") = True Then
        '            '1.5.7
        '            cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam"))
        '            MyConnObj.Execute cm
        '            ''''''
        '            myrecfam.Delete
        '        Else
        '            'asignación de indice
        '            MyRecAux.Open "select * from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
        '            If MyRecAux.EOF Then
        '                MyRecAux.AddNew
        '                MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
        '                MyRecAux.Fields("familyindex") = 999999
        '                MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
        '                MyRecAux.Update
        '            End If
        '            MyRecAux.Close
        '            Set MyRecAux = Nothing
        '
        '            Set MyRecFam1 = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(myrecfam.Fields("codi_fam")) & " and index<>0")
        '            If Not MyRecFam1.EOF Then
        '                'If lUpperNW Then
        '                '    cm = "update pcscale.counterfamilies set familyindex=" & CStr(MyRecFam1.Fields("index") Mod (MyRecFam.Fields("secc_maqui") * 100)) & " where familycode=" & CStr(MyRecFam.Fields("codi_fam"))
        '                'Else
        '                    cm = "update pcscale.counterfamilies set familyindex=" & CStr(MyRecFam1.Fields("index")) & " where familycode=" & CStr(myrecfam.Fields("codi_fam"))
        '                'End If
        '                MyConnObj.Execute cm
        '            End If
        '
        '            MyRecFam1.Close
        '            Set MyRecFam1 = Nothing
        '
        '        End If
        '        If Not myrecfam.EOF Then myrecfam.Movenext
        '    Loop
        'End If
        'myrecfam.Close
        'Set myrecfam = Nothing
        
        Set myrecfam = Mybase.OpenRecordset("select * from familias")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
        If Not myrecfam.EOF Then
            myrecfam.MoveFirst
            Do While Not myrecfam.EOF
                If myrecfam.Fields("borrado") = True Or myrecfam.Fields("index") = 0 Then
                    '1.5.7
                    cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                    MyConnObj.Execute cm
                    ''''''
                    myrecfam.Delete
                Else
                    'asignación de indice
                    cm = "delete from pcscale.counterfamilies where familyindex=" & CStr(myrecfam.Fields("index")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                    MyConnObj.Execute cm
                    cm = "delete from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui"))
                    MyConnObj.Execute cm
                    
                    MyRecAux.Open "select * from pcscale.counterfamilies where familycode=" & CStr(myrecfam.Fields("codi_fam")) & " and countercode=" & CStr(myrecfam.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                        MyRecAux.Fields("countercode") = myrecfam.Fields("secc_maqui")
                        MyRecAux.Fields("familyindex") = myrecfam.Fields("index")
                        MyRecAux.Fields("familycode") = myrecfam.Fields("codi_fam")
                        MyRecAux.Update
                    End If
                    MyRecAux.Close
                    Set MyRecAux = Nothing
        
                End If
                If Not myrecfam.EOF Then myrecfam.Movenext
            Loop
        End If
        myrecfam.Close
        Set myrecfam = Nothing
    
    End If
    
'CadenadeLog "Fin familias 'familias' a counterfamilies..."

End If

    lEsAlta = False
    
    '''''''''''''''''''''''''
    Set MyRecVar = Mybase.OpenRecordset("select * from articulo where borrado=true order by codigo")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo CargaCadena(1409) & " " & CStr(MyRecVar.Fields("codigo"))
            
            cm = "delete from pcscale.items where code=" & CStr(MyRecVar.Fields("codigo"))
            MyConnObj.Execute cm
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo"))
            MyConnObj.Execute cm
            cm = "delete from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error Resume Next
            cm = "delete from text15 where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error GoTo 0
            On Error Resume Next
            cm = "delete from teclas where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute cm
            On Error GoTo 0
            Do_Events
            '3.8.1
            If lFornes Then
                'Public Function Item_OrderInFamily_T(ByVal nFam As Long, ByVal nMethod As Long) As Long
                nFI = Item_OrderInFamily_T(MyRecVar.Fields("codi_fam"), 3)
            End If
            ''''''
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    
    If lEsTest Then
        'Baja especial teclas...
        Set MyRecVar = Mybase.OpenRecordset("select * from teclas where borrado=true order by codigo")
        If Not MyRecVar.EOF Then
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecVar.Recordcount
            MyRecVar.MoveFirst
            Do While Not MyRecVar.EOF
                frmMonitor.MostrarInfo "Baja especial Teclas..." & " " & CStr(MyRecVar.Fields("codigo"))
    
                cm = "delete from pcscale.items where code=" & CStr(MyRecVar.Fields("codigo"))
                MyConnObj.Execute cm
                cm = "update articulo set tran_plu=' ' where codigo=" & CStr(MyRecVar.Fields("codigo"))
                Mybase.Execute cm
    
                Do_Events
    
                MyRecVar.Delete
                If Not MyRecVar.EOF Then MyRecVar.Movenext
            Loop
    
        End If
        MyRecVar.Close
        Set MyRecVar = Nothing
    End If
    
    MyRecAux.Open "SELECT plu,counter FROM pcscale.items where plu<>0 GROUP BY plu,counter HAVING COUNT(*)>1", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecAux.EOF Then
        MyRecAux.MoveFirst
        Do While Not MyRecAux.EOF
            MyConnObj.Execute "delete from pcscale.items where counter=" & CStr(MyRecAux.Fields("counter")) & " and plu=" & CStr(MyRecAux.Fields("plu"))
            MyRecAux.Movenext
        Loop
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo where tran_plu<>'*' order by codigo")
        lEnvImage = False
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo order by codigo")
        lEnvImage = True
    End If
    
    If Not MyRecVar.EOF Then
        MyRecVar.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
            
        nContEvnt = 0
        frmMonitor.MostrarInfo CargaCadena(230) & " " & CStr(frmMonitor.barra1.Max)
        CadenadeLog "Artículos a enviar...:" & CStr(frmMonitor.barra1.Max)
        Do While Not MyRecVar.EOF
            
            
            lEsAlta = False
            On Error Resume Next
            
            MyRecAux.Open "select * from pcscale.items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
'If MyRecVar.Fields("codigo") = 31431 Then
'lEsAlta = lEsAlta
'End If

            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            Else
                'musgrave
                'MyConnObj.Execute "delete from pcscale.items where code=" & MyRecVar.Fields("codigo")
                'MyRecAux.AddNew
                'lEsAlta = True
                ''''''''''''
            End If
            
            MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
            If MyRecVar.Fields("codigo") = 30 Then
                lEsAlta = lEsAlta
            End If
            If lEsAlta Or lEnvShort Then
                MyRecAux.Fields("shortname") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu1")))
            End If
            MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu1")))
            
            '''''''''''''''''3.2.5 se anula
            'If lEsAlta And lEsTest Then
            '    If bbDDver < 80 Then
            '        MyRecAux.Fields("onkey") = 0
            '        MyRecAux.Fields("familyorder") = 0 '1000000 + MyRecVar.Fields("codigo")
            '    Else
            '        MyRecAux.Fields("onkey") = 0
            '        MyRecAux.Fields("familyorder") = 0 '1000000 + MyRecVar.Fields("codigo")
            '    End If
            'Else
            '    If lEsAlta = False Then
            '        If MyRecAux.Fields("familyorder") > 1000000 And bbDDver >= 80 Then
            '            MyRecAux.Fields("familyorder") = 0
            '            MyRecAux.Fields("onkey") = 1
            '        End If
            '    Else
            '        If bbDDver < 80 Then
            '            MyRecAux.Fields("onkey") = 0
            '            MyRecAux.Fields("familyorder") = 0 '1000000 + MyRecVar.Fields("codigo")
            '        Else
            '            MyRecAux.Fields("onkey") = 0
            '            MyRecAux.Fields("familyorder") = 0 '1000000 + MyRecVar.Fields("codigo")
            '        End If
            '    End If
            'End If
            '''''''''''''''''
            If lEsTest Then MyRecAux.Fields("onkey") = 1
            '''''''''''''''''
'CadenadeLog ".1"
            If lUpper = False And lAgora = False Then
            
                MyRecAux.Fields("text") = ""
                MyRecAux.Fields("text1") = ""
                MyRecAux.Fields("text2") = ""
                MyRecAux.Fields("text3") = ""
                MyRecAux.Fields("text4") = ""
                MyRecAux.Fields("text5") = ""
'CadenadeLog ".2"
                cm1 = ""
                cm = ""
                'If Dir(App.Path & "\40l.txt") = "" Then '1.6.9
                If l40TXT Then
                    
                    If lTransPlu Then
                        For ncont = 1 To 20
                            Select Case ncont
                            Case 1
                                cm = cm & MyRecVar.Fields("des_plu2")
                            Case 2
                                cm = cm & MyRecVar.Fields("des_plu3")
                            Case 3
                                cm = cm & MyRecVar.Fields("des_plu4")
                            Case 4
                                cm = cm & MyRecVar.Fields("des_plu5")
                            Case 5
                                cm = cm & MyRecVar.Fields("des_plu6")
                            Case 6
                                cm = cm & MyRecVar.Fields("des_plu7")
                            Case 7
                                cm = cm & MyRecVar.Fields("des_plu8")
                            Case 8
                                cm = cm & MyRecVar.Fields("des_plu9")
                            Case 9
                                cm = cm & MyRecVar.Fields("des_plu0")
                            Case 10
                                cm1 = cm1 & MyRecVar.Fields("des_plux")
                            Case 11
                                cm1 = cm1 & MyRecVar.Fields("des_plu11")
                            Case 12
                                cm1 = cm1 & MyRecVar.Fields("des_plu12")
                            Case 13
                                cm1 = cm1 & MyRecVar.Fields("des_plu13")
                            Case 14
                                cm1 = cm1 & MyRecVar.Fields("des_plu14")
                            Case 15
                                cm1 = cm1 & MyRecVar.Fields("des_plu15")
                            Case 16
                                cm1 = cm1 & MyRecVar.Fields("des_plu16")
                            Case 17
                                cm1 = cm1 & MyRecVar.Fields("des_plu17")
                            Case 18
                                cm1 = cm1 & MyRecVar.Fields("des_plu18")
                            Case 19
                                cm1 = cm1 & MyRecVar.Fields("des_plu19")
                            Case 20
                                cm1 = cm1 & MyRecVar.Fields("des_plu20")
                        
                            End Select
                        
                        Next ncont
                    Else
                        For ncont = 1 To 20
                            Select Case ncont
                            Case 1
                                cm = cm & MyRecVar.Fields("des_plu2")
                            Case 2
                                cm = cm & MyRecVar.Fields("des_plu3")
                            Case 3
                                cm = cm & MyRecVar.Fields("des_plu4")
                            Case 4
                                cm = cm & MyRecVar.Fields("des_plu5")
                            Case 5
                                cm = cm & MyRecVar.Fields("des_plu6")
                            Case 6
                                cm = cm & MyRecVar.Fields("des_plu7")
                            Case 7
                                cm = cm & MyRecVar.Fields("des_plu8")
                            Case 8
                                cm = cm & MyRecVar.Fields("des_plu9")
                            Case 9
                                cm = cm & MyRecVar.Fields("des_plu0")
                            Case 10
                                cm = cm & MyRecVar.Fields("des_plux")
                            Case 11
                                cm = cm & MyRecVar.Fields("des_plu11")
                            Case 12
                                cm = cm & MyRecVar.Fields("des_plu12")
                            Case 13
                                cm = cm & MyRecVar.Fields("des_plu13")
                            Case 14
                                cm = cm & MyRecVar.Fields("des_plu14")
                            Case 15
                                cm = cm & MyRecVar.Fields("des_plu15")
                            Case 16
                                cm = cm & MyRecVar.Fields("des_plu16")
                            Case 17
                                cm = cm & MyRecVar.Fields("des_plu17")
                            Case 18
                                cm = cm & MyRecVar.Fields("des_plu18")
                            Case 19
                                cm = cm & MyRecVar.Fields("des_plu19")
                            Case 20
                                cm = cm & MyRecVar.Fields("des_plu20")
                        
                            End Select
                        
                        Next ncont
                    
                    End If
                    
'If MyRecVar.Fields("codigo") = 70281 Then
'cm = cm
'End If
'CadenadeLog ".3"
                    Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                    If Not myRecTmp1.EOF Then
                        If lTransPlu Then
                            For ncont = 21 To 80
                                cm1 = cm1 & Mid(myRecTmp1.Fields("txt_" & CStr(ncont)), 1, 24)
                            Next ncont
                        Else
                            For ncont = 21 To 40
                                cm1 = cm1 & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
                            Next ncont
                        End If
                    End If
                    myRecTmp1.Close
                    Set myRecTmp1 = Nothing
                    If lTransPlu Then
                        cm1 = verLF(cm1)
                        cm = verLF(cm)
                        If Trim(cm1) <> "" Or Trim(cm) <> "" Then
CadenadeLog "1-->cm1=" & cm1
CadenadeLog "1-->cm=" & cm

                            MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm1)
                            MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(cm)
                            MyRecAux.Fields("text2") = ""
                            MyRecAux.Fields("text3") = ""
                            MyRecAux.Fields("text4") = ""
                            MyRecAux.Fields("text5") = ""
                        Else
                            'If lEsAlta Then
                            MyRecAux.Fields("text") = ""
                            MyRecAux.Fields("text1") = ""
                            MyRecAux.Fields("text2") = ""
                            MyRecAux.Fields("text3") = ""
                            MyRecAux.Fields("text4") = ""
                            MyRecAux.Fields("text5") = ""
                        End If
                    Else
                        'cm1 = verLF(cm & cm1)
                        'If Trim(cm1) <> "" Then
                        '    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm1) '1.6.9 se recortaba a 1024
                        'Else
                        '    'If lEsAlta Then
                        '    MyRecAux.Fields("text") = ""
                        'End If
                        If l40TXT Then
                            If Trim(cm & cm1) <> "" Then
CadenadeLog "2-->cm1=" & cm1
CadenadeLog "2-->cm=" & cm
'For ncont = 1 To Len(cm1)
'    CadenadeLog Format(ncont, "000") & "-->" & Mid(cm1, ncont, 1) & "-->" & CStr(Asc(Mid(cm1, ncont, 1)))'
'Next ncont
                                If l2140Text1 Then
                                    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm))
                                    MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(verLF(cm1)) '1.6.9 se recortaba a 1024
                                Else
If Trim(cm) = Trim(cm1) Then
                                    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm)) '1.6.9 se recortaba a 1024
Else
                                    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm & cm1)) '1.6.9 se recortaba a 1024
End If
                                End If
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text") = ""
                            End If
                        End If
                        
                    End If
                Else
                    
                    If l39TXT Then
                        cm = ""
                        cm1 = ""
'''
                        For ncont = 1 To 19
                            Select Case ncont
                            Case 1
                                cm = cm & MyRecVar.Fields("des_plu3")
                            Case 2
                                cm = cm & MyRecVar.Fields("des_plu4")
                            Case 3
                                cm = cm & MyRecVar.Fields("des_plu5")
                            Case 4
                                cm = cm & MyRecVar.Fields("des_plu6")
                            Case 5
                                cm = cm & MyRecVar.Fields("des_plu7")
                            Case 6
                                cm = cm & MyRecVar.Fields("des_plu8")
                            Case 7
                                cm = cm & MyRecVar.Fields("des_plu9")
                            Case 8
                                cm = cm & MyRecVar.Fields("des_plu0")
                            Case 9
                                cm = cm & MyRecVar.Fields("des_plux")
                            Case 10
                                cm1 = cm1 & MyRecVar.Fields("des_plu11")
                            Case 11
                                cm1 = cm1 & MyRecVar.Fields("des_plu12")
                            Case 12
                                cm1 = cm1 & MyRecVar.Fields("des_plu13")
                            Case 13
                                cm1 = cm1 & MyRecVar.Fields("des_plu14")
                            Case 14
                                cm1 = cm1 & MyRecVar.Fields("des_plu15")
                            Case 15
                                cm1 = cm1 & MyRecVar.Fields("des_plu16")
                            Case 16
                                cm1 = cm1 & MyRecVar.Fields("des_plu17")
                            Case 17
                                cm1 = cm1 & MyRecVar.Fields("des_plu18")
                            Case 18
                                cm1 = cm1 & MyRecVar.Fields("des_plu19")
                            Case 19
                                cm1 = cm1 & MyRecVar.Fields("des_plu20")
                            End Select
                        
                        Next ncont
                    
                        Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                        If Not myRecTmp1.EOF Then
                            For ncont = 21 To 40
                                cm1 = cm1 & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
                            Next ncont
                        End If
                        myRecTmp1.Close
                        Set myRecTmp1 = Nothing

                        cm = verLF(cm & cm1)
                        MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm)
                        MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(MyRecVar.Fields("des_plu2"))

CadenadeLog "3-->cm1=" & cm1
CadenadeLog "3-->cm=" & cm

'''
                    Else
                        cm = ""
                        cm1 = ""
                        If lTransPlu Then
                            For ncont = 1 To 20
                                Select Case ncont
                                Case 1
                                    cm = cm & MyRecVar.Fields("des_plu2")
                                Case 2
                                    cm = cm & MyRecVar.Fields("des_plu3")
                                Case 3
                                    cm = cm & MyRecVar.Fields("des_plu4")
                                Case 4
                                    cm = cm & MyRecVar.Fields("des_plu5")
                                Case 5
                                    cm = cm & MyRecVar.Fields("des_plu6")
                                Case 6
                                    cm = cm & MyRecVar.Fields("des_plu7")
                                Case 7
                                    cm = cm & MyRecVar.Fields("des_plu8")
                                Case 8
                                    cm = cm & MyRecVar.Fields("des_plu9")
                                Case 9
                                    cm = cm & MyRecVar.Fields("des_plu0")
                                Case 10
                                    cm1 = cm1 & MyRecVar.Fields("des_plux")
                                Case 11
                                    cm1 = cm1 & MyRecVar.Fields("des_plu11")
                                Case 12
                                    cm1 = cm1 & MyRecVar.Fields("des_plu12")
                                Case 13
                                    cm1 = cm1 & MyRecVar.Fields("des_plu13")
                                Case 14
                                    cm1 = cm1 & MyRecVar.Fields("des_plu14")
                                Case 15
                                    cm1 = cm1 & MyRecVar.Fields("des_plu15")
                                Case 16
                                    cm1 = cm1 & MyRecVar.Fields("des_plu16")
                                Case 17
                                    cm1 = cm1 & MyRecVar.Fields("des_plu17")
                                Case 18
                                    cm1 = cm1 & MyRecVar.Fields("des_plu18")
                                Case 19
                                    cm1 = cm1 & MyRecVar.Fields("des_plu19")
                                Case 20
                                    cm1 = cm1 & MyRecVar.Fields("des_plu20")
                            
                                End Select
                            
                            Next ncont
                        
                        Else
                            If l5l Then
                                cm = ""
                                cm1 = ""
                                For ncont = 6 To 20
                                    Select Case ncont
                                    Case 6
                                        cm = cm & MyRecVar.Fields("des_plu7")
                                    Case 7
                                        cm = cm & MyRecVar.Fields("des_plu8")
                                    Case 8
                                        cm = cm & MyRecVar.Fields("des_plu9")
                                    Case 9
                                        cm = cm & MyRecVar.Fields("des_plu0")
                                    Case 10
                                        cm = cm & MyRecVar.Fields("des_plux")
                                    Case 11
                                        cm = cm & MyRecVar.Fields("des_plu11")
                                    Case 12
                                        cm = cm & MyRecVar.Fields("des_plu12")
                                    Case 13
                                        cm = cm & MyRecVar.Fields("des_plu13")
                                    Case 14
                                        cm = cm & MyRecVar.Fields("des_plu14")
                                    Case 15
                                        cm = cm & MyRecVar.Fields("des_plu15")
                                    Case 16
                                        cm = cm & MyRecVar.Fields("des_plu16")
                                    Case 17
                                        cm = cm & MyRecVar.Fields("des_plu17")
                                    Case 18
                                        cm = cm & MyRecVar.Fields("des_plu18")
                                    Case 19
                                        cm = cm & MyRecVar.Fields("des_plu19")
                                    Case 20
                                        cm = cm & MyRecVar.Fields("des_plu20")
                                
                                    End Select
                                
                                Next ncont
                            
                            
                            End If
                        End If
'CadenadeLog ".4"
                        Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                        If Not myRecTmp1.EOF Then
                            
                            If lTransPlu Then
                                For ncont = 21 To 80
                                    cm1 = cm1 & Mid(myRecTmp1.Fields("txt_" & CStr(ncont)), 1, 24)
                                Next ncont
                            Else
                                For ncont = 21 To 40
                                    cm1 = cm1 & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
                                Next ncont
                            End If
                        End If
                        myRecTmp1.Close
                        Set myRecTmp1 = Nothing
                        If lTransPlu Then
                            cm1 = verLF(cm1)
                            cm = verLF(cm)
                            
CadenadeLog "4-->cm1=" & cm1
CadenadeLog "4-->cm=" & cm
                            
                            If Trim(cm1) <> "" Or Trim(cm) <> "" Then
                                MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm1)
                                MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(cm)
                                MyRecAux.Fields("text2") = ""
                                MyRecAux.Fields("text3") = ""
                                MyRecAux.Fields("text4") = ""
                                MyRecAux.Fields("text5") = ""
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text") = ""
                                MyRecAux.Fields("text1") = ""
                                MyRecAux.Fields("text2") = ""
                                MyRecAux.Fields("text3") = ""
                                MyRecAux.Fields("text4") = ""
                                MyRecAux.Fields("text5") = ""
                            End If
                        End If
                    
                    End If
                    
                End If
'CadenadeLog ".5"
                'cm = ""
                'If l5l And lTransPlu = False And l39TXT = False Then
                '
                '    For ncont = 1 To 20
                '        Select Case ncont
                '        Case 1
                '            cm = cm & Trim(MyRecVar.Fields("des_plu2"))
                '        Case 2
                '            cm = cm & Trim(MyRecVar.Fields("des_plu3"))
                '        Case 3
                '            cm = cm & Trim(MyRecVar.Fields("des_plu4"))
                '        Case 4
                '            cm = cm & Trim(MyRecVar.Fields("des_plu5"))
                '        Case 5
                '            cm = cm & Trim(MyRecVar.Fields("des_plu6"))
                '        Case 6
                '            cm = cm & Trim(MyRecVar.Fields("des_plu7"))
                '        Case 7
                '            cm = cm & Trim(MyRecVar.Fields("des_plu8"))
                '        Case 8
                '            cm = cm & Trim(MyRecVar.Fields("des_plu9"))
                '        Case 9
                '            cm = cm & Trim(MyRecVar.Fields("des_plu0"))
                '        Case 10
                '            cm = cm & Trim(MyRecVar.Fields("des_plux"))
                '        Case 11
                '            cm = cm & Trim(MyRecVar.Fields("des_plu11"))
                '        Case 12
                '            cm = cm & Trim(MyRecVar.Fields("des_plu12"))
                '        Case 13
                '            cm = cm & Trim(MyRecVar.Fields("des_plu13"))
                '        Case 14
                '            cm = cm & Trim(MyRecVar.Fields("des_plu14"))
                '        Case 15
                '            cm = cm & Trim(MyRecVar.Fields("des_plu15"))
                '        Case 16
                '            cm = cm & Trim(MyRecVar.Fields("des_plu16"))
                '        Case 17
                '            cm = cm & Trim(MyRecVar.Fields("des_plu17"))
                '        Case 18
                '            cm = cm & Trim(MyRecVar.Fields("des_plu18"))
                '        Case 19
                '            cm = cm & Trim(MyRecVar.Fields("des_plu19"))
                '        Case 20
                '            cm = cm & Trim(MyRecVar.Fields("des_plu20"))
                '
                '        End Select
                '
                '    Next ncont
                '
                '    'If Dir(App.Path & "\40l.txt") <> "" Then '1.6.9
                    'If l40TXT Then
                    '    If Trim(cm & cm1) <> "" Then
                    '        If l2140Text1 Then
                    '            MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm))
                    '            MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(verLF(cm1)) '1.6.9 se recortaba a 1024
                    '        Else
                    '            MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm & cm1)) '1.6.9 se recortaba a 1024
                    '        End If
                    '    Else
                    '        'If lEsAlta Then
                    '        MyRecAux.Fields("text") = ""
                    '    End If
                    'End If
                '        s(1) = Mid(cm, 1, 100)
                '        s(2) = Mid(cm, 101, 100)
                '        s(3) = Mid(cm, 201, 100)
                '        s(4) = Mid(cm, 301, 100)
                '        s(5) = Mid(cm, 401)
                '        If Len(s(1)) > 0 Then
                '            s(1) = verLF(s(1))
                '        End If
                '        If Len(s(2)) > 0 Then
                '            s(2) = verLF(s(2))
                '        End If
                '        If Len(s(3)) > 0 Then
                '            s(3) = verLF(s(3))
                '        End If
                '        If Len(s(4)) > 0 Then
                '            s(4) = verLF(s(4))
                '        End If
                '        If Len(s(5)) > 0 Then
                '            s(5) = verLF(s(5))
                '        End If
                '
                        If l5l = False Then
                '            If Trim(s(1)) <> "" Then
                '                MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(Trim(s(1)))
                '            Else
                '                'If lEsAlta Then
                '                MyRecAux.Fields("text1") = ""
                '            End If
                '            If Trim(s(2)) <> "" Then
                '                MyRecAux.Fields("text2") = ConvertStringToUtf8Bytes(Trim(s(2)))
                '            Else
                '                'If lEsAlta Then
                '                MyRecAux.Fields("text2") = ""
                '            End If
                '            If Trim(s(3)) <> "" Then
                '                MyRecAux.Fields("text3") = ConvertStringToUtf8Bytes(Trim(s(3)))
                '            Else
                '                'If lEsAlta Then
                '                MyRecAux.Fields("text3") = ""
                '            End If
                '            If Trim(s(4)) <> "" Then
                '                MyRecAux.Fields("text4") = ConvertStringToUtf8Bytes(Trim(s(4)))
                '            Else
                '                'If lEsAlta Then
                '                MyRecAux.Fields("text4") = ""
                '            End If
                '            If Trim(s(5)) <> "" Then
                '                MyRecAux.Fields("text5") = ConvertStringToUtf8Bytes(Trim(s(5)))
                '            Else
                '                'If lEsAlta Then
                '                MyRecAux.Fields("text5") = ""
                '            End If
                '
                        Else
                            If Trim(MyRecVar.Fields("des_plu2")) <> "" Then
                                MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu2")))
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text1") = ""
                            End If
                            If Trim(MyRecVar.Fields("des_plu3")) <> "" Then
                                MyRecAux.Fields("text2") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu3")))
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text2") = ""
                            End If
                            If Trim(MyRecVar.Fields("des_plu4")) <> "" Then
                                MyRecAux.Fields("text3") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu4")))
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text3") = ""
                            End If
                            If Trim(MyRecVar.Fields("des_plu5")) <> "" Then
                                MyRecAux.Fields("text4") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu5")))
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text4") = ""
                            End If
                            If Trim(MyRecVar.Fields("des_plu6")) <> "" Then
                                MyRecAux.Fields("text5") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu6")))
                            Else
                                'If lEsAlta Then
                                MyRecAux.Fields("text5") = ""
                            End If
                             MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm & cm1))
                        End If
                        'MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(verLF(cm1))
                '    End If
                    If l40TXT Then
                        cm1 = verLF(cm & cm1)
                        MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm1)
                        cm1 = ""
                        cm = ""
                    End If
                    
                    'If lNewTEXT And l40TXT Then
                        Set myRecTmp1 = Mybase.OpenRecordset("select * from text15 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                        If Not myRecTmp1.EOF Then
                            If Len(Trim(myRecTmp1.Fields("text1"))) <> 0 Then MyRecAux.Fields("text1") = verLF(ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text1"))))
                            If Len(Trim(myRecTmp1.Fields("text1"))) <> 0 Then MyRecAux.Fields("text") = verLF(ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text1"))))
          CadenadeLog "es descritivo 1"
                            If Len(Trim(myRecTmp1.Fields("text2"))) <> 0 Then MyRecAux.Fields("text2") = ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text2")))
                            If Len(Trim(myRecTmp1.Fields("text3"))) <> 0 Then MyRecAux.Fields("text3") = ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text3")))
                            If Len(Trim(myRecTmp1.Fields("text4"))) <> 0 Then MyRecAux.Fields("text4") = ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text4")))
                            If Len(Trim(myRecTmp1.Fields("text5"))) <> 0 Then MyRecAux.Fields("text5") = ConvertStringToUtf8Bytes(Trim(myRecTmp1.Fields("text5")))
                            
                            'If Len(Trim(myRecTmp1.Fields("text1"))) <> 0 Or Len(Trim(myRecTmp1.Fields("text2"))) <> 0 Or Len(Trim(myRecTmp1.Fields("text3"))) <> 0 Or _
                            'Len(Trim(myRecTmp1.Fields("text4"))) <> 0 Or Len(Trim(myRecTmp1.Fields("text5"))) <> 0 Then
                            '    cm1 = verLF(cm & cm1)
                            '    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(cm1)
                            '    cm1 = ""
                            '    cm = ""
                            'End If
                            
                        End If
                        myRecTmp1.Close
                        Set myRecTmp1 = Nothing
                    'End If
                'Else
                '
                '    If lTransPlu = False And l5l = True Then
                '
                '        cm = Trim(MyRecVar.Fields("des_plu2"))
                '        cm = verLF(cm)
                '        If Trim(cm) <> "" Then
                '            MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(Trim(cm))
                '        Else
                '            'If lEsAlta Then
                '            MyRecAux.Fields("text1") = ""
                '        End If
                '        cm = Trim(MyRecVar.Fields("des_plu3"))
                '        cm = verLF(cm)
                '        If Trim(cm) <> "" Then
                '            MyRecAux.Fields("text2") = ConvertStringToUtf8Bytes(Trim(cm))
                '        Else
                '            'If lEsAlta Then
                '            MyRecAux.Fields("text2") = ""
                '        End If
                '        cm = Trim(MyRecVar.Fields("des_plu4"))
                '        cm = verLF(cm)
                '        If Trim(cm) <> "" Then
                '            MyRecAux.Fields("text3") = ConvertStringToUtf8Bytes(Trim(cm))
                '        Else
                '            'If lEsAlta Then
                '            MyRecAux.Fields("text3") = ""
                '        End If
                '        cm = Trim(MyRecVar.Fields("des_plu5"))
                '        cm = verLF(cm)
                '        If Trim(cm) <> "" Then
                '            MyRecAux.Fields("text4") = ConvertStringToUtf8Bytes(Trim(cm))
                '        Else
                '            'If lEsAlta Then
                '            MyRecAux.Fields("text4") = ""
                '        End If
                '        cm = Trim(MyRecVar.Fields("des_plu6"))
                '        cm = verLF(cm)
                '        If Trim(cm) <> "" Then
                '            MyRecAux.Fields("text5") = ConvertStringToUtf8Bytes(Trim(cm))
                '        Else
                '            'If lEsAlta Then
                '            MyRecAux.Fields("text5") = ""
                '        End If
                '        cm = verLF(Trim(MyRecVar.Fields("des_plu7"))) & verLF(Trim(MyRecVar.Fields("des_plu8"))) & verLF(Trim(MyRecVar.Fields("des_plu9")))
                '        cm = cm & verLF(Trim(MyRecVar.Fields("des_plu0"))) & verLF(Trim(MyRecVar.Fields("des_plux"))) & verLF(Trim(MyRecVar.Fields("des_plu11")))
                '        cm = cm & verLF(Trim(MyRecVar.Fields("des_plu12"))) & verLF(Trim(MyRecVar.Fields("des_plu13"))) & verLF(Trim(MyRecVar.Fields("des_plu14")))
                '        cm = cm & verLF(Trim(MyRecVar.Fields("des_plu15"))) & verLF(Trim(MyRecVar.Fields("des_plu16"))) & verLF(Trim(MyRecVar.Fields("des_plu17")))
                '        cm = cm & verLF(Trim(MyRecVar.Fields("des_plu18"))) & verLF(Trim(MyRecVar.Fields("des_plu19"))) & verLF(Trim(MyRecVar.Fields("des_plu20")))
                '        cm1 = cm & cm1
                '        MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(Trim(cm1))
                '    End If
               '
                'End If
                
            Else
             
                'cm = ""
                'For ncont = 1 To 20
                ' Select Case ncont
                ' Case 1
                '     cm = cm & Trim(MyRecVar.Fields("des_plu2"))
                ' Case 2
                '     cm = cm & Trim(MyRecVar.Fields("des_plu3"))
                ' Case 3
                '     cm = cm & Trim(MyRecVar.Fields("des_plu4"))
                ' Case 4
                '     cm = cm & Trim(MyRecVar.Fields("des_plu5"))
                ' Case 5
                '     cm = cm & Trim(MyRecVar.Fields("des_plu6"))
                ' Case 6
                '     cm = cm & Trim(MyRecVar.Fields("des_plu7"))
                ' Case 7
                '     cm = cm & Trim(MyRecVar.Fields("des_plu8"))
                ' Case 8
                '     cm = cm & Trim(MyRecVar.Fields("des_plu9"))
                ' Case 9
                '     cm = cm & Trim(MyRecVar.Fields("des_plu0"))
                ' Case 10
                '     cm = cm & Trim(MyRecVar.Fields("des_plux"))
                ' Case 11
                '     cm = cm & Trim(MyRecVar.Fields("des_plu11"))
                ' Case 12
                '     cm = cm & Trim(MyRecVar.Fields("des_plu12"))
                ' Case 13
                '     cm = cm & Trim(MyRecVar.Fields("des_plu13"))
                ' Case 14
                '     cm = cm & Trim(MyRecVar.Fields("des_plu14"))
                ' Case 15
                '     cm = cm & Trim(MyRecVar.Fields("des_plu15"))
                ' Case 16
                '     cm = cm & Trim(MyRecVar.Fields("des_plu16"))
                ' Case 17
                '     cm = cm & Trim(MyRecVar.Fields("des_plu17"))
                ' Case 18
                '     cm = cm & Trim(MyRecVar.Fields("des_plu18"))
                ' Case 19
                '     cm = cm & Trim(MyRecVar.Fields("des_plu19"))
                ' Case 20
                '     cm = cm & Trim(MyRecVar.Fields("des_plu20"))
                '
                ' End Select
                'Next ncont
             
                'Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                'If Not myRecTmp1.EOF Then
                '    For ncont = 21 To 40
                '        cm = cm & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
                '    Next ncont
                'End If
                'myRecTmp1.Close
                'Set myRecTmp1 = Nothing
                '
                'If Trim(cm) <> "" Then
                '    cm = ConvertStringToUtf8Bytes(verLF(cm))
                '    MyRecAux.Fields("text") = Mid(cm, 1, 1024)
                'Else
                '    If lEsAlta Then MyRecAux.Fields("text") = ""
                'End If
                cm = ""
                Set myRecTmp1 = Mybase.OpenRecordset("select * from text15 where codigo=" & CStr(MyRecVar.Fields("codigo")))
                If Not myRecTmp1.EOF Then
                    cm = Trim(myRecTmp1.Fields("text1"))
                End If
                myRecTmp1.Close
                Set myRecTmp1 = Nothing
                
                If Trim(cm) <> "" Then
                    cm = ConvertStringToUtf8Bytes(verLF(cm))
                    MyRecAux.Fields("text") = Mid(cm, 1, 1024)
                    MyRecAux.Fields("text1") = Mid(cm, 1, 1024)
                Else
                    If lEsAlta Then MyRecAux.Fields("text") = ""
                    If lEsAlta Then MyRecAux.Fields("text1") = ""
                End If
                
                'If lEsAlta Then MyRecAux.Fields("text1") = ""
                If lEsAlta Then MyRecAux.Fields("text2") = ""
                If lEsAlta Then MyRecAux.Fields("text3") = ""
                If lEsAlta Then MyRecAux.Fields("text4") = ""
                If lEsAlta Then MyRecAux.Fields("text5") = ""
            
            End If
'CadenadeLog ".6"
            CadenadeLog "Código:" & CStr(MyRecVar.Fields("codigo")) & " Precio:" & CStr(MyRecVar.Fields("precio")) & " Euros:" & CStr(MyRecVar.Fields("euros"))
            
            If Dir(App.Path & "\continuos.bat") <> "" Then
                frmMonitor.MostrarInfo "Código:" & CStr(MyRecVar.Fields("codigo")) & " Precio:" & CStr(MyRecVar.Fields("precio")) & " Euros:" & CStr(MyRecVar.Fields("euros"))
            End If
            
            MyRecAux.Fields("family") = MyRecVar.Fields("codi_fam")
            If UsaEuro Then
                MyRecAux.Fields("price") = MyRecVar.Fields("euros")
            Else
                MyRecAux.Fields("price") = MyRecVar.Fields("precio")
            End If
            If MyRecVar.Fields("codi_pes") = "W" Then
                MyRecAux.Fields("saleform") = 1
            Else
                MyRecAux.Fields("saleform") = 0
            End If
            MyRecAux.Fields("plu") = MyRecVar.Fields("plu")
            MyRecAux.Fields("counter") = MyRecVar.Fields("secc_maqui")
            '3.8.1
            'MyRecAux.Fields("section") = MyRecVar.Fields("secc_maqui")
            ''''''
            
CadenadeLog "Cuenca Art."

            If Trim(MyRecVar.Fields("tran_ep")) <> "" Then
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("tran_ep")))
            Else
                
                Select Case MyRecVar.Fields("etq")
                Case 0
                    MyRecAux.Fields("label") = ""
                Case 16
                    MyRecAux.Fields("label") = "-"
                Case Else
                    MyRecAux.Fields("label") = ConvertStringToUtf8Bytes(CStr(MyRecVar.Fields("etq")) & ".xml")
                End Select
                
            End If
            If Trim(MyRecVar.Fields("label2")) <> "" Then
                MyRecAux.Fields("label2") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("label2")))
            Else
                MyRecAux.Fields("label2") = "-"
            End If

            If nTaxFam(MyRecVar.Fields("codi_fam")) = 0 Then
                MyRecAux.Fields("vat") = MyRecVar.Fields("tipo_iva")
            Else
                MyRecAux.Fields("vat") = nTaxFam(MyRecVar.Fields("codi_fam"))
            End If
            If sLabel1(MyRecVar.Fields("codi_fam")) <> "" Then
                MyRecAux.Fields("label") = quita_ext(sLabel1(MyRecVar.Fields("codi_fam")))
            Else
                'MyRecAux.Fields("label") = ""
            End If
            If slabel2(MyRecVar.Fields("codi_fam")) <> "" Then
                MyRecAux.Fields("label2") = quita_ext(slabel2(MyRecVar.Fields("codi_fam")))
            Else
                'MyRecAux.Fields("label2") = ""
            End If
            
            'If lEsAlta Then
                MyRecAux.Fields("discount") = MyRecVar.Fields("ning")
                If lEsAlta Then MyRecAux.Fields("qlevel") = ""
            'End If
            ''''''''''''
            'If bbDDver < 103 Then
'CadenadeLog ".7"
            If IsNull(MyRecVar.Fields("art_cb")) Then 'If IsNull(MyRecAux.Fields("ean13")) Then
                MyRecAux.Fields("ean13") = ""
            Else
                sXcb = MyRecVar.Fields("art_cb")
                sXcb = Replace(sXcb, ".", "")
                sXcb = Replace(sXcb, ",", "")
                sXcb = Replace(sXcb, Chr(133), "")
                sXcb = Trim(sXcb)
                If IsNumeric(sXcb) Then
                    If Trim(MyRecVar.Fields("art_cb")) <> "" Then
                        MyRecAux.Fields("ean13") = ConvertStringToUtf8Bytes(sXcb)
                    Else
                        MyRecAux.Fields("ean13") = ""
                    End If
                Else
                    MyRecAux.Fields("ean13") = ""
                End If
            End If
'CadenadeLog ".9"
            If IsNull(MyRecVar.Fields("art_cb")) Then
                MyRecAux.Fields("ean13print") = ""
            Else
                If Trim(MyRecVar.Fields("art_cb")) <> "" Then
                    MyRecAux.Fields("ean13print") = Mid(ConvertStringToUtf8Bytes(sXcb), 1, 12)
                Else
                    MyRecAux.Fields("ean13print") = ""
                End If
            End If
'CadenadeLog ".10"
            '''''''''''
            'josue
            If MyRecVar.Fields("posicion") <> 0 Then
                MyRecAux.Fields("nutrition") = MyRecVar.Fields("posicion")
            End If
            '''''''''''
            If lAgora = False Or lEsAlta Then
                MyRecAux.Fields("expiring") = MyRecVar.Fields("caducidad")
                MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
                MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
            Else
                'If lEsAlta Then
                '    MyRecAux.Fields("expiring") = MyRecVar.Fields("caducidad")
                '    MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
                '    MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
                'Else
                    If MyRecVar.Fields("caducidad") <> 0 Then
                        MyRecAux.Fields("expiring") = MyRecVar.Fields("caducidad")
                    End If
                    If MyRecVar.Fields("pref") <> 0 Then
                        MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
                    End If
                    If (MyRecVar.Fields("tara") / 1000) <> 0 Then
                        MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
                    End If
                'End If
            End If
            'MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
            'MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
'CadenadeLog ".11"
            '" _+_ "
            If Not IsNull(MyRecVar.Fields("imagen")) Then
                s(1) = separaFich(Trim(MyRecVar.Fields("imagen")))
                If Not IsNull(s(1)) Then
                    If Trim(s(1)) <> "" Then
                        MyRecAux.Fields("icon") = Trim(ConvertStringToUtf8Bytes(s(1)))
                    Else
                        If lEsAlta Then MyRecAux.Fields("icon") = ""
                    End If
                Else
                    If lEsAlta Then MyRecAux.Fields("icon") = ""
                End If
            Else
                If lEsAlta Then MyRecAux.Fields("icon") = ""
            End If
            If MyRecVar.Fields("lmix") Then
                MyRecAux.Fields("mixed") = 1
            Else
                MyRecAux.Fields("mixed") = 0
            End If
            MyRecAux.Fields("mixedweight") = MyRecVar.Fields("poid") / 1000
            '''''''''''''
            If MyRecVar.Fields("onkey") = 0 Then
                MyRecAux.Fields("onkey") = 0
            End If
            '''''''''''''3.8.1
            If lFornes Then
                MyRecAux.Fields("section") = MyRecVar.Fields("secc_maqui")
            End If
            ''''''
                
            
'CadenadeLog ".12"
            MyRecAux.Update
            
            '3.8.1
            If lFornes And lEsAlta Then
                nFI = Item_OrderInFamily_T(MyRecVar.Fields("codi_fam"), 3)
            End If
            ''''''
''''
'DROP TABLE IF EXISTS `pcscale`.`itemtextbonus`;
'CREATE TABLE  `pcscale`.`itemtextbonus` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) NOT NULL,
'  `NumLine` smallint(6) DEFAULT '1',
'  `Text` varchar(40) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `itemtext_idx` (`Code`,`NumLine`),
'  CONSTRAINT `itemtextbonus_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
'REPLACE into table (id, name, age) values(1, "A", 19)
''''Restablecer
            On Error GoTo finNcont
            'For ncont = 1 To 20
            '    Select Case ncont
            '    Case 1
            '        'If Trim(MyRecVar.Fields("des_plu2")) <> "" Then
            '            MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu2")), 1, 40) & "')"
            '        'End If
            '    Case 2
            '        'If Trim(MyRecVar.Fields("des_plu3")) <> "" Then
            '            MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu3")), 1, 40) & "')"
            '        'End If'
            '
            '    Case 3
            '        'If Trim(MyRecVar.Fields("des_plu4")) <> "" Then
            '            MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu4")), 1, 40) & "')"
            '        'End If
            '
            '    Case 4
            '        'If Trim(MyRecVar.Fields("des_plu5")) <> "" Then
            '            MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu5")), 1, 40) & "')"
            '        'End If
'
'                Case 5
'                    'If Trim(MyRecVar.Fields("des_plu6")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu6")), 1, 40) & "')"
'                    'End If''
'
'                Case 6
'                    'If Trim(MyRecVar.Fields("des_plu7")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu7")), 1, 40) & "')"
'                    'End If''
'
'                Case 7
'                    'If Trim(MyRecVar.Fields("des_plu8")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu8")), 1, 40) & "')"
'                    'End If''
'
'                Case 8
'                    'If Trim(MyRecVar.Fields("des_plu9")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu9")), 1, 40) & "')"
'                    'End If'
'
'                Case 9
'                    'If Trim(MyRecVar.Fields("des_plu0")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu0")), 1, 40) & "')"
'                    'End If'
'
'                Case 10
'                    'If Trim(MyRecVar.Fields("des_plux")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plux")), 1, 40) & "')"
'                    'End If'
'
'                Case 11
'                    'If Trim(MyRecVar.Fields("des_plu11")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu11")), 1, 40) & "')"
'                    'End If''
'
'                Case 12
'                    'If Trim(MyRecVar.Fields("des_plu12")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu12")), 1, 40) & "')"
'                    'End If'
'
'                Case 13
'                    'If Trim(MyRecVar.Fields("des_plu13")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu13")), 1, 40) & "')"
'                    'End If''
'
'                Case 14
'                    'If Trim(MyRecVar.Fields("des_plu14")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu14")), 1, 40) & "')"
'                    'End If''
'
'                Case 15
'                    'If Trim(MyRecVar.Fields("des_plu15")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu15")), 1, 40) & "')"
'                    'End If'
''
'                Case 16
'                    'If Trim(MyRecVar.Fields("des_plu16")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu16")), 1, 40) & "')"
'                    'End If'
'
'                Case 17
'                    'If Trim(MyRecVar.Fields("des_plu17")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu17")), 1, 40) & "')"
'                    'End If''
'
'                Case 18
'                    'If Trim(MyRecVar.Fields("des_plu18")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu18")), 1, 40) & "')"
'                    'End If'
'
'                Case 19
'                    'If Trim(MyRecVar.Fields("des_plu19")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu19")), 1, 40) & "')"
'                    'End If'
'
'                Case 20
'                    'If Trim(MyRecVar.Fields("des_plu20")) <> "" Then
'                        MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(MyRecVar.Fields("des_plu20")), 1, 40) & "')"
'                    'End If''
'
'                End Select
'            Next ncont
'            Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
'            If Not myRecTmp1.EOF Then
'                For ncont = 21 To 40
'                    MyConnObj.Execute "replace into pcscale.itemtextbonus (code,numline,text) values (" & CStr(MyRecVar.Fields("codigo")) & "," & CStr(ncont) & ",'" & Mid(Trim(myRecTmp1.Fields("txt_" & CStr(ncont))), 1, 40) & "')"
'                Next ncont
'            End If
            'myRecTmp1.Close
            'Set myRecTmp1 = Nothing
'''''''''''''''''''''''''''''
'CadenadeLog ".13"
finNcont:
On Error GoTo 0
On Error GoTo FINARTICULO
'''''
            
            MyRecAux.Close
            Set MyRecAux = Nothing

''''''''''''''
            If MyRecVar.Fields("usatramos") Then
                On Error GoTo SigueTec
                Set MyRecSec = Mybase.OpenRecordset("select * from tramos where codigo=" & CStr(MyRecVar.Fields("codigo")))
                If Not MyRecSec.EOF Then
                    If MyRecSec.Fields("borrado") Then
                        cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=0"
                        MyConnObj.Execute cm
                        MyRecSec.Delete
                    Else
                        cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=0"
                        MyConnObj.Execute cm
                        MyRecAux.Open "select * from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=0", MyConnObj, adOpenKeyset, adLockOptimistic
                        If UsaEuro Then
                            If MyRecSec.Fields("euros1") <> 0 And MyRecSec.Fields("tramo1") <> 0 Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
                                MyRecAux.Fields("rate") = 0
                                MyRecAux.Fields("price") = MyRecSec.Fields("euros1")
                                MyRecAux.Fields("step") = MyRecSec.Fields("tramo1")
                                MyRecAux.Update
                            End If
                            If MyRecSec.Fields("euros2") <> 0 And MyRecSec.Fields("tramo2") <> 0 Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
                                MyRecAux.Fields("rate") = 0
                                MyRecAux.Fields("price") = MyRecSec.Fields("euros2")
                                MyRecAux.Fields("step") = MyRecSec.Fields("tramo2")
                                MyRecAux.Update
                            End If
                        Else
                            If MyRecSec.Fields("precio1") <> 0 And MyRecSec.Fields("tramo1") <> 0 Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
                                MyRecAux.Fields("rate") = 0
                                MyRecAux.Fields("price") = MyRecSec.Fields("precio1")
                                MyRecAux.Fields("step") = MyRecSec.Fields("tramo1")
                                MyRecAux.Update
                            End If
                            If MyRecSec.Fields("precio2") <> 0 And MyRecSec.Fields("tramo2") <> 0 Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
                                MyRecAux.Fields("rate") = 0
                                MyRecAux.Fields("price") = MyRecSec.Fields("precio2")
                                MyRecAux.Fields("step") = MyRecSec.Fields("tramo2")
                                MyRecAux.Update
                            End If
                        End If

                        MyRecAux.Close
                        Set MyRecAux = Nothing
                    End If
                Else
                    cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=0"
                    MyConnObj.Execute cm
                End If
                MyRecSec.Close
                Set MyRecSec = Nothing
                cm = "update tramos set tran_tramo='*' where codigo=" & CStr(MyRecVar.Fields("codigo"))
                Mybase.Execute cm
            Else
                cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=0"
                MyConnObj.Execute cm
            End If
            On Error GoTo SigueTec
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=1"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar1") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",1," & Quitar_Coma(CStr(MyRecVar.Fields("tar1")), Len(CStr(MyRecVar.Fields("tar1")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=2"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar2") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",2," & Quitar_Coma(CStr(MyRecVar.Fields("tar2")), Len(CStr(MyRecVar.Fields("tar2")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=3"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar3") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",3," & Quitar_Coma(CStr(MyRecVar.Fields("tar3")), Len(CStr(MyRecVar.Fields("tar3")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=4"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar4") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",4," & Quitar_Coma(CStr(MyRecVar.Fields("tar4")), Len(CStr(MyRecVar.Fields("tar4")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=5"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar5") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",5," & Quitar_Coma(CStr(MyRecVar.Fields("tar5")), Len(CStr(MyRecVar.Fields("tar5")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=6"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar6") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",6," & Quitar_Coma(CStr(MyRecVar.Fields("tar6")), Len(CStr(MyRecVar.Fields("tar6")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=7"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar7") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",7," & Quitar_Coma(CStr(MyRecVar.Fields("tar7")), Len(CStr(MyRecVar.Fields("tar7")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=8"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar8") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",8," & Quitar_Coma(CStr(MyRecVar.Fields("tar8")), Len(CStr(MyRecVar.Fields("tar8")))) & ",0)"
                MyConnObj.Execute cm
            End If
            cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=9"
            MyConnObj.Execute cm
            If MyRecVar.Fields("tar9") <> 0 Then
                cm = "insert into pcscale.itempricerates (code,rate,price,step) values (" & CStr(MyRecVar.Fields("codigo")) & ",9," & Quitar_Coma(CStr(MyRecVar.Fields("tar9")), Len(CStr(MyRecVar.Fields("tar9")))) & ",0)"
                MyConnObj.Execute cm
            End If
            'cm = "delete from pcscale.itempricerates where code=" & CStr(MyRecVar.Fields("codigo")) & " and rate=10"
            'MyConnObj.Execute cm
            'If MyRecVar.Fields("tar10") <> 0 Then
            '    cm = "insert into pcscale.itempricerates ('code','rate','price','step') values (" & CStr(MyRecVar.Fields("codigo")) & ",10," & CStr(MyRecVar.Fields("tar5")) & ",0)"
            '    MyConnObj.Execute cm
            'End If
            
'CadenadeLog ".14"

SigueTec:
On Error GoTo 0
On Error GoTo FINARTICULO

''''''''''''''''''

            MyRecVar.Edit
            MyRecVar.Fields("tran_plu") = "*"
            'MyRecVar.Fields("tran_tramo") = "*"
            MyRecVar.Update
            
            MyRecVar.Movenext
            
            'Call MonitorBarra

            'Call frmControl.Balanza_ComOK
            
            'Do_Events '--> ver
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 3 Then
                Call MonitorBarra
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirArt
            
        Loop
        
    End If
    
''''''''''''TECLAS...
'DROP TABLE IF EXISTS `pcscale`.`plukeys`;
'CREATE TABLE  `pcscale`.`plukeys` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Section` int(11) DEFAULT NULL,
'  `Board` int(11) DEFAULT NULL,
'  `PLUKey` int(11) DEFAULT NULL,
'  `Code` int(11) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Section` (`Section`,`Board`,`PLUKey`),
'  KEY `Code` (`Code`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
    
    If lEsTest = False Then
    
        frmMonitor.MostrarInfo CargaCadena(234)  '"ENVÍO DE TECLAS RÁPIDAS"
        Dim nEqp As Integer
            
        
        Set MyRecSec = Mybase.OpenRecordset("select * from seccion order by codi_ident")
        If Not MyRecSec.EOF Then
            MyRecSec.MoveFirst
            Do While Not MyRecSec.EOF
                
                Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and (borrado=true or codigo=0) and codi_ident=" & CStr(MyRecSec.Fields("codi_ident")))
                If Not myRecTmp1.EOF Then
            
                    myRecTmp1.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = myRecTmp1.Recordcount
                    myRecTmp1.MoveFirst
                    Do While Not myRecTmp1.EOF
                        
                        '4.6.0
                        'If IsNull(myRecTmp1.Fields("numero_eqp")) Then
                        '    nEqp = 0
                        'Else
                            nEqp = Val(myRecTmp1.Fields("tran_tecsc10")) - 1 'guardar en texto GRUPO (1 a 9)
                        'End If
                        '''''''
                        
                        cm = "delete from pcscale.plukeys where section=" & CStr(nEqp + 1) & " and board=" & CStr(myRecTmp1.Fields("tabla")) & " and plukey=" & CStr(myRecTmp1.Fields("codi_tec"))
                        MyConnObj.Execute cm
                        
                        ''myRecTmp1.Delete
                        'If Not myRecTmp1.EOF Then myRecTmp1.Movenext
                    
                        Call MonitorBarra
                    
                        'Do_Events '--> ver
                        nContEvnt = nContEvnt + 1
                        If nContEvnt >= 20 Then
                            nContEvnt = 0
                            DoEvents
                            Sleep (1)
                        End If
                        
                        '''''''''''''''''''
                    
                        'myRecTmp1.Delete
                        If Not myRecTmp1.EOF Then myRecTmp1.Movenext
                    
                    Loop
                
               
                End If
                myRecTmp1.Close
                Set myRecTmp1 = Nothing
                
                Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and (borrado=true or codigo=0) and codi_ident=" & CStr(MyRecSec.Fields("codi_ident")))
                If Not myRecTmp1.EOF Then
            
                    myRecTmp1.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = myRecTmp1.Recordcount
                    myRecTmp1.MoveFirst
                    Do While Not myRecTmp1.EOF
                        
                        myRecTmp1.Delete
                        If Not myRecTmp1.EOF Then myRecTmp1.Movenext
                    
                    Loop
                
               
                End If
                myRecTmp1.Close
                Set myRecTmp1 = Nothing
                
                
                If Modificaciones Then
                    Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=" & CStr(MyRecSec.Fields("codi_ident")) & " and ((tran_tec<>'*') or isnull(tran_tec))")
                Else
                    Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=" & CStr(MyRecSec.Fields("codi_ident"))) ' where codigo=" & CStr(MyRecVar.Fields("codigo")))
                End If
                If Not myRecTmp1.EOF Then
                    myRecTmp1.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = myRecTmp1.Recordcount
                    myRecTmp1.MoveFirst
                    Do While Not myRecTmp1.EOF
                        'If IsNull(myRecTmp1.Fields("numero_eqp")) Then
                        '    nEqp = 0
                        'Else
                            nEqp = myRecTmp1.Fields("tran_tecsc10") - 1
                        'End If
                        MyRecAux.Open "select * from pcscale.plukeys where section=" & CStr(nEqp + 1) & " and board=" & CStr(myRecTmp1.Fields("tabla")) & " and plukey=" & CStr(myRecTmp1.Fields("codi_tec")), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            'Call frmControl.Balanza_ComError(0)
                            GoTo FINARTICULO
                        End If
                        On Error GoTo FINARTICULO
                        lEsAlta = False
                        If MyRecAux.EOF Then
                            MyRecAux.AddNew
                            lEsAlta = True
                        End If
                        
                        MyRecAux.Fields("section") = nEqp + 1
                        MyRecAux.Fields("board") = myRecTmp1.Fields("tabla")
                        MyRecAux.Fields("plukey") = myRecTmp1.Fields("codi_tec")
                        MyRecAux.Fields("code") = myRecTmp1.Fields("codigo")
                        
                        MyRecAux.Update
                        MyRecAux.Close
                        Set MyRecAux = Nothing
                        myRecTmp1.Edit
                        'If myRecTmp1.Fields("tran_tec") = "E" Then
                        '    myRecTmp1.Fields("tran_tec") = "A"
                        'Else
                        '    myRecTmp1.Fields("tran_tec") = "T"
                        'End If
                        myRecTmp1.Fields("tran_tec") = "*"
                        myRecTmp1.Update
                        myRecTmp1.Movenext
                    
                        Call MonitorBarra
                    
                        'Do_Events '--> ver
                        nContEvnt = nContEvnt + 1
                        If nContEvnt >= 20 Then
                            nContEvnt = 0
                            DoEvents
                            Sleep (1)
                        End If
                        '''''''''''''''''''
                    
                    Loop
                
                End If
                myRecTmp1.Close
                Set myRecTmp1 = Nothing
    
                MyRecSec.Movenext
            Loop
        End If
        '''aquí
                If Modificaciones Then
                    Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=0" & " and ((tran_tec<>'*') or isnull(tran_tec))")
                Else
                    Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=0")  ' where codigo=" & CStr(MyRecVar.Fields("codigo")))
                End If
                If Not myRecTmp1.EOF Then
                    myRecTmp1.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = myRecTmp1.Recordcount
                    myRecTmp1.MoveFirst
                    Do While Not myRecTmp1.EOF
                        'If IsNull(myRecTmp1.Fields("numero_eqp")) Then
                        '    nEqp = 0
                        'Else
                            nEqp = myRecTmp1.Fields("tran_tecsc10") - 1
                        'End If
                        MyRecAux.Open "select * from pcscale.plukeys where section=" & CStr(nEqp + 1) & " and board=" & CStr(myRecTmp1.Fields("tabla")) & " and plukey=" & CStr(myRecTmp1.Fields("codi_tec")), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            'Call frmControl.Balanza_ComError(0)
                            GoTo FINARTICULO
                        End If
                        On Error GoTo FINARTICULO
                        lEsAlta = False
                        If MyRecAux.EOF Then
                            MyRecAux.AddNew
                            lEsAlta = True
                        End If
                        
                        MyRecAux.Fields("section") = nEqp + 1
                        MyRecAux.Fields("board") = myRecTmp1.Fields("tabla")
                        MyRecAux.Fields("plukey") = myRecTmp1.Fields("codi_tec")
                        MyRecAux.Fields("code") = myRecTmp1.Fields("codigo")
                        
                        MyRecAux.Update
                        MyRecAux.Close
                        Set MyRecAux = Nothing
                        myRecTmp1.Edit
                        'If myRecTmp1.Fields("tran_tec") = "E" Then
                        '    myRecTmp1.Fields("tran_tec") = "A"
                        'Else
                        '    myRecTmp1.Fields("tran_tec") = "T"
                        'End If
                        myRecTmp1.Fields("tran_tec") = "*"
                        myRecTmp1.Update
                        myRecTmp1.Movenext
                    
                        Call MonitorBarra
                    
                        'Do_Events '--> ver
                        nContEvnt = nContEvnt + 1
                        If nContEvnt >= 20 Then
                            nContEvnt = 0
                            DoEvents
                            Sleep (1)
                        End If
                        '''''''''''''''''''
                    
                    Loop
                
                End If
                myRecTmp1.Close
                Set myRecTmp1 = Nothing
    
        '''''''
    Else
        nContEvnt = 0
        If Modificaciones Then
            Set MyRecVar = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and ((tran_tec<>'*') or isnull(tran_tec))")
        Else
            Set MyRecVar = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0")
        End If
        If Not MyRecVar.EOF Then
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecVar.Recordcount
            MyRecVar.MoveFirst
            Do While Not MyRecVar.EOF
                frmMonitor.MostrarInfo "Programación Especial Tecla Código:" & " " & CStr(MyRecVar.Fields("codigo"))
                
                MyRecAux.Open "select * from pcscale.items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
                
                If Not MyRecAux.EOF Then
                    If MyRecVar.Fields("tabla") <> 0 Then
                        MyRecAux.Fields("family") = MyRecVar.Fields("tabla")
                        MyRecAux.Fields("onkey") = 1
                        MyRecAux.Fields("familyorder") = MyRecVar.Fields("codi_tec")
                        MyRecAux.Update
                    End If
                    'Do_Events
                    nContEvnt = nContEvnt + 1
                    If nContEvnt >= 20 Then
                        nContEvnt = 0
                        DoEvents
                        Sleep (1)
                    End If
                    
                End If
                MyRecAux.Update
                
                MyRecAux.Close
                Set MyRecAux = Nothing
                
                MyRecVar.Edit
                MyRecVar.Fields("tran_tec") = "*"
                MyRecVar.Update
                If Not MyRecVar.EOF Then MyRecVar.Movenext
            Loop
        
        End If
        MyRecVar.Close
        Set MyRecVar = Nothing
    
    End If
'''''''''''''''''''''''''''''''
             
SalirArt:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    
CadenadeLog "Fin envio Artículos."
    
    MUE_EnviaArticuloNW = True

'QUITAR raspberry
'GoTo raspberrY
'CadenadeLog "regresa a Raspberry..."
'''''''''''''''''
    
    Exit Function
FINARTICULO:
    CadenadeLog "Error en MUE_EnviaArticuloNW:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
    
    
    MUE_EnviaArticuloNW = False
End Function

'Adaptar TouchScale Barna
'DROP TABLE IF EXISTS `pcscale`.`bovine`;
'CREATE TABLE  `pcscale`.`bovine` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumCard` smallint(3) NOT NULL,
'  `EarTag` varchar(30) NOT NULL,
'  `BirthLand` smallint(6) DEFAULT NULL,
'  `BirthDate` date DEFAULT NULL,
'  `BreedingLand` smallint(6) DEFAULT NULL,
'  `SlaughterLand` smallint(6) DEFAULT NULL,
'  `SlaughterReg` varchar(30) DEFAULT NULL,
'  `SlaughterDate` date DEFAULT NULL,
'  `ButcherLand` smallint(6) DEFAULT NULL,
'  `ButcherReg` varchar(30) DEFAULT NULL,
'  `Lot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumCard` (`NumCard`)
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
Public Sub MUE_EnviarFichasVacuno(ByVal NSistema As Integer, ByVal Modificaciones As Boolean, ByVal Marcar As Boolean, Optional MyCod As Long)
'Dim Crianza As Long
'Dim Nacimiento As Long
'Dim Produccion As Long
'Dim sCrianza As String
'Dim sNacimiento As String
'Dim sProduccion As Long
'Dim Sacrificio As Long
'Dim sSacrificio As String
'Dim Despiece As Long
'Dim sDespiece As Long
'Dim sSexo As String
'Dim sRaza As String
'Dim sCategoria As String
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nC1 As Integer
Dim cm As String
'Dim nLast As Double

    If MyCod = 0 Then frmMonitor.MostrarInfo CargaCadena(619) '"Envío de Fichas Vacuno"
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
        Do_Events
        If MyCod = 0 Then
            If Modificaciones Then
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno where  (tranficha<>'*' OR tranficha is null)")
            Else
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno ")
            End If
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from fichavacuno where borrado=false and codigo=" & Chr(34) & MyCod & Chr(34))
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    If !borrado = False Then
                        'Crianza = !codcrianza
                        'On Error Resume Next
                        'sCrianza = Trim(!valorpcrianza)
                        'Nacimiento = !codnacimiento
                        'sNacimiento = Trim(!valorpnacimiento)
                        'Produccion = !codproduccion
                        'sProduccion = Trim(!valorpproduccion)
                        'Sacrificio = !codpsacrificio
                        'sSacrificio = Trim(!valorpsacrificio)
                        'Despiece = !codpdespiece
                        'sDespiece = Trim(!valorpdespiece)
                        'On Error GoTo 0
                        'If !codcrianza = 0 And !codnacimiento = 0 Then
                        '    Crianza = Produccion
                        '    sCrianza = sProduccion
                        '    Nacimiento = Produccion
                        '    sNacimiento = sProduccion
                        'End If
                        'Valores paises, categoria, raza, sexo
                        
                        ''''''''''''''''''''''''''''''''''''''
                        MyRecDes.Open "Select * from pcscale.bovine where numcard=" & CStr(!codigo), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            If MyRecDes.EOF Then

                                MyRecDes.AddNew
                                MyRecDes.Fields("numcard") = !codigo
                                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(!Identificador))
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(!rsisacrificio))
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(!rsidespiece))
                                
                                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(!lote))
                                    
                            
                            Else
                                   
                                MyRecDes.Fields("numcard") = !codigo
                                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(!Identificador))
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(!rsisacrificio))
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(!rsidespiece))
                                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(!lote))
                                   
                               
                            End If
                            
                            MyRecDes.Update
                            MyRecDes.Close
                            Set MyRecDes = Nothing
                            
                            frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo, "000") & " --> OK"
                            'Edit_Record Registro
                            .Edit
                            .Fields("tranficha") = "*"
                            .Update
    
                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        cm = "delete from pcscale.bovine where numcard=" & CStr(!codigo)
                        MyConnObj.Execute cm
                        ''''''''''''''''''''
                        frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tranficha") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirFicha:
                    
                Loop
            End If
        End With
    
SalirFicha:
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing

salir:

End Sub

'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `TIN` varchar(16) NOT NULL DEFAULT '',
'  `Name` varchar(255) NOT NULL,
'  `Address` varchar(255) DEFAULT NULL,
'  `City` varchar(255) DEFAULT NULL,
'  `WEB` varchar(255) DEFAULT NULL,
'  `Phone` varchar(255) DEFAULT NULL,
'  `Text` varchar(255) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `TIN` (`TIN`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;

'>= 72
'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `TIN` varchar(16) NOT NULL DEFAULT '',
'  `Name` varchar(255) NOT NULL DEFAULT '',
'  `Address` varchar(255) DEFAULT '',
'  `City` varchar(255) DEFAULT '',
'  `Postalcode` varchar(80) DEFAULT '',
'  `Country` varchar(80) DEFAULT '',
'  `WEB` varchar(255) DEFAULT '',
'  `Phone` varchar(80) DEFAULT '',
'  `Fax` varchar(80) DEFAULT '',
'  `MobilePhone` varchar(80) DEFAULT '',
'  `eMail` varchar(80) DEFAULT '',
'  `Text` varchar(255) DEFAULT '',
'  `Discount` decimal(4,2) DEFAULT '0.00',
'  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `TIN` (`TIN`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaCLientes(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean)
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
'Dim MyRecTmp As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim cm As String

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCLIENTES
    End If
    On Error GoTo 0
    On Error GoTo FINCLIENTES
    Set MyRecVar = Mybase.OpenRecordset("select * from cliente where borrado=true")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo CargaCadena(1410) & CStr(MyRecVar.Fields("cod_cli"))
            'Cm = "delete pcscale.customers where tin=" & CStr(MyRecVar.Fields("cod_cli"))
            cm = "delete from pcscale.customers where id=" & CStr(MyRecVar.Fields("cod_cli"))
            MyConnObj.Execute cm
            cm = "delete from pcscale.customers where trim(upper(tin))='" & Trim(UCase(MyRecVar.Fields("data3")) & "'")
            MyConnObj.Execute cm
            
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente where tran_cli<>'*'")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            
            'MyRecDes.Open "Select * from pcscale.customers where tin=" & CStr(MyRecVar.Fields("cod_cli")), MyConnObj, adOpenKeyset, adLockOptimistic
            'MyRecDes.Open "Select * from pcscale.customers where id=" & CStr(MyRecVar.Fields("cod_cli")), MyConnObj, adOpenKeyset, adLockOptimistic
            MyRecDes.Open "Select * from pcscale.customers where trim(upper(tin))='" & Trim(UCase(MyRecVar.Fields("data3")) & "'"), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINCLIENTES
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            
            frmMonitor.MostrarInfo CargaCadena(1411) & " " & CStr(MyRecVar.Fields("cod_cli"))
'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
 ' `TIN` varchar(16) NOT NULL DEFAULT '',
 ' `Name` varchar(255) NOT NULL DEFAULT '',
'  `Address` varchar(255) DEFAULT '',
'  `City` varchar(255) DEFAULT '',
'  `Postalcode` varchar(80) DEFAULT '',
'  `Country` varchar(80) DEFAULT '',
'  `WEB` varchar(255) DEFAULT '',
'  `Phone` varchar(80) DEFAULT '',
'  `Fax` varchar(80) DEFAULT '',
'  `MobilePhone` varchar(80) DEFAULT '',
'  `eMail` varchar(80) DEFAULT '',
'  `Text` varchar(255) DEFAULT '',
'  `Discount` decimal(4,2) DEFAULT '0.00',
'  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `TIN` (`TIN`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=23457 DEFAULT CHARSET=utf8;

            'MyRecDes.Fields("tin") = MyRecVar.Fields("cod_cli")
            MyRecDes.Fields("id") = MyRecVar.Fields("cod_cli")
            'If MyRecVar.Fields("cod_cli") = 2222 Then
            'bbDDver = bbDDver
            'End If
            MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nom_cli")))
            MyRecDes.Fields("address") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("data1")))
            MyRecDes.Fields("city") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data2")))
            MyRecDes.Fields("tin") = Mid(ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data3"))), 1, 16)
            MyRecDes.Fields("phone") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data4")))
            MyRecDes.Fields("text") = ""
            If bbDDver >= 72 Then
'  `Discount` decimal(4,2) DEFAULT '0.00',
'  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
                MyRecDes.Fields("Discount") = MyRecVar.Fields("discount")
                If MyRecVar.Fields("factura") Then
                    MyRecDes.Fields("SimplifiedInvoice") = 1
                Else
                    MyRecDes.Fields("SimplifiedInvoice") = 0
                End If
            End If
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_cli") = "*"
            MyRecVar.Update
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCliente
            
        Loop
    
    
    End If
SalirCliente:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCLIENTES:
    CadenadeLog "Error en MUE_EnviaClientes:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
End Sub

Public Function Ii(ByVal sS As String) As String
Dim nPos As Integer
Dim sR As String
    sR = sS
    nPos = InStr(1, sR, "I")
    If nPos <> 0 Then
        Do While InStr(1, sR, "I") <> 0
            nPos = InStr(1, sR, "I")
            sR = Mid(sR, 1, nPos - 1) & "i" & Mid(sR, nPos + 1)
        Loop
    End If
    Ii = sR
End Function

'Adaptar TouchScale Barna
Public Sub MUE_EnviaCodigosBarras(NSistema As Integer, Modificaciones As Boolean, Optional Marcar As Boolean)
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecTmp As dao.Recordset
'Dim MyRecDes As New ADODB.Recordset
Dim MyRecAux As New ADODB.Recordset
Dim cm As String

    frmMonitor.MostrarInfo CargaCadena(233)
    
    Set Mybase = dao.OpenDatabase(Base_General)
    
        Do_Events
        If Modificaciones Then
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar where (tran_bar<>'*' OR tran_bar is null)")
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar")
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    
                    If !borrado = False Then
                        If !numero_eqp <> 0 Then
                            ' Equipo
                            CadenadeLog "Consulta Equipo " & CStr(!numero_eqp)
                            MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(!numero_eqp), MyConnObj, adOpenKeyset, adLockOptimistic
                        Else
                            ' Seccion
                            CadenadeLog "Consulta Seccion " & CStr(!secc_maqui)
                            MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(!secc_maqui), MyConnObj, adOpenKeyset, adLockOptimistic
                        End If
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            
                            If MyRecAux.EOF Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("localsales") = ""
                                MyRecAux.Fields("localLineSLB") = ""
                                MyRecAux.Fields("localTotalSLB") = ""
                                MyRecAux.Fields("localSS") = ""
                                MyRecAux.Fields("localPacket1") = ""
                                MyRecAux.Fields("localPacket2") = ""
                                MyRecAux.Fields("localtotal") = ""
                                MyRecAux.Fields("networksales") = ""
                                MyRecAux.Fields("NetworkLineSLB") = ""
                                MyRecAux.Fields("NetworkTotalSLB") = ""
                                MyRecAux.Fields("NetworkSS") = ""
                                MyRecAux.Fields("NetworkPacket1") = ""
                                MyRecAux.Fields("NetworkPacket2") = ""
                                CadenadeLog "Nuevo item"
                            End If
                            
                            If IsNull(!numero_eqp) Then
                                ' Seccion
                                MyRecAux.Fields("number") = !secc_maqui
                                MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("C")
                            Else
                                If !numero_eqp <> 0 Then
                                    ' Balanza
                                    MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("S")
                                    MyRecAux.Fields("number") = !numero_eqp
                                Else
                                    ' Seccion
                                    MyRecAux.Fields("number") = !secc_maqui
                                    MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("C")
                                End If
                            End If
                            If !secc_eqp = 0 Then
                                ' Red
'  `NetworkSales` char(12) NOT NULL,
'  `NetworkParked` char(12) NOT NULL DEFAULT '',
'  `NetworkLineSLB` char(12) NOT NULL,
'  `NetworkTotalSLB` char(12) NOT NULL,
'  `NetworkSS` char(12) NOT NULL,
'  `NetworkPacket1` char(12) NOT NULL,
'  `NetworkPacket2` char(12) NOT NULL,
'  `LocalSales` char(12) NOT NULL,
'  `LocalLineSLB` char(12) NOT NULL,
'  `LocalTotalSLB` char(12) NOT NULL,
'  `LocalSS` char(12) NOT NULL,
'  `LocalPacket1` char(12) NOT NULL,
'  `LocalPacket2` char(12) NOT NULL,
'  `LocalTotal` char(12) NOT NULL,
                                MyRecAux.Fields("networksales") = ConvertStringToUtf8Bytes(Trim(!CB_VNT))
                                MyRecAux.Fields("NetworkParked") = ConvertStringToUtf8Bytes(Trim(!CB_VNT))
                                MyRecAux.Fields("NetworkLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_SUP)))
                                MyRecAux.Fields("NetworkTotalSLB") = ConvertStringToUtf8Bytes((Trim(!cb_mix))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_MIX)))
                                If Len(Trim(!CB_EV1)) <= 12 Then
                                    MyRecAux.Fields("NetworkSS") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                Else
                                    'MyRecAux.Fields("NetworkSS") = "25CCCCCiiiii"
                                End If
                                If Len(Trim(!CB_EV1)) <= 12 Then
                                    MyRecAux.Fields("NetworkPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                Else
                                    'MyRecAux.Fields("NetworkPacket1") = "25CCCCCiiiii"
                                    MyConnObj.Execute "INSERT INTO code128 (counter, code1) VALUES (" & CStr(!secc_maqui) & ",'" & Trim(!CB_EV1) & "') ON DUPLICATE KEY UPDATE code1='" & Trim(!CB_EV1) & "'"
                                End If
                                If Len(Trim(!CB_EV2)) <= 12 Then
                                    MyRecAux.Fields("NetworkPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV2)))
                                Else
                                    'MyRecAux.Fields("NetworkPacket2") = "25CCCCCiiiii"
                                    MyConnObj.Execute "INSERT INTO code128 (counter, code2) VALUES (" & CStr(!secc_maqui) & ",'" & Trim(!CB_EV2) & "') ON DUPLICATE KEY UPDATE code2='" & Trim(!CB_EV2) & "'"
                                End If
                            Else
                                ' Local
                                MyRecAux.Fields("localsales") = ConvertStringToUtf8Bytes(Trim(!CB_VNT))
                                MyRecAux.Fields("localLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_SUP)))
                                MyRecAux.Fields("localTotalSLB") = ConvertStringToUtf8Bytes((Trim(!cb_mix))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_MIX)))
                                If Len(Trim(!CB_EV1)) <= 12 Then
                                    MyRecAux.Fields("localSS") = ConvertStringToUtf8Bytes(Trim(Ii(!CB_EV1)))
                                Else
                                    'MyRecAux.Fields("localSS") = "25CCCCCiiiii"
                                End If
                                If Len(Trim(!CB_EV1)) <= 12 Then
                                    MyRecAux.Fields("localPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                Else
                                    'MyRecAux.Fields("localPacket1") = "25CCCCCiiiii"
                                    MyConnObj.Execute "INSERT INTO code128 (counter, code1) VALUES (" & CStr(!secc_maqui) & ",'" & Trim(!CB_EV1) & "') ON DUPLICATE KEY UPDATE code1='" & Trim(!CB_EV1) & "'"
                                End If
                                If Len(Trim(!CB_EV2)) <= 12 Then
                                    MyRecAux.Fields("localPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV2)))
                                Else
                                    'MyRecAux.Fields("localPacket2") = "25CCCCCiiiii"
                                    MyConnObj.Execute "INSERT INTO code128 (counter, code2) VALUES (" & CStr(!secc_maqui) & ",'" & Trim(!CB_EV2) & "') ON DUPLICATE KEY UPDATE code2='" & Trim(!CB_EV2) & "'"
                                End If
                                MyRecAux.Fields("localtotal") = ConvertStringToUtf8Bytes((Trim(!cb_mix))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                            End If
                               
                            
                            MyRecAux.Update
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            If IsNull(!numero_eqp) Then
                                frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                            Else
                                If !numero_eqp Then
                                    frmMonitor.MostrarDato CargaCadena(23) & " EQP.:" & Format(!numero_eqp, "000") & " --> OK"
                                Else
                                    frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                                End If
                            End If
                            .Edit
                            .Fields("tran_bar") = "*"
                            .Update

                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        If !numero_eqp <> 0 Then
                            cm = "delete from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(!numero_eqp)
                        Else
                            cm = "delete from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(!secc_maqui)
                        End If
                        
                        MyConnObj.Execute cm
                        ''''''''''''''''''''
                        'frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tran_bar") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirBarras
                    
                Loop
            End If
        End With
    
SalirBarras:
    MyRecVar.Close
    Set MyRecVar = Nothing
salir:
    Mybase.Close
    Set Mybase = Nothing

End Sub


Public Sub MUE_EnviaVendedor(NSistema As Integer, Modificaciones As Boolean)
    '///////////////////////////////////////
    '// Envío de vendedores (TouchScale)    //
    '///////////////////////////////////////
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim cm As String
Dim lEsAlta As Boolean
Dim sColorTmp As String
Dim nColor As Integer

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINVENDEDOR
    End If
    On Error GoTo 0
    On Error GoTo FINVENDEDOR
    
    'Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where borrado=true or tran_vend<>'*'")
    Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where borrado=true")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            'frmMonitor.MostrarInfo "Baja Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
            CadenadeLog "Baja Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
            cm = "delete from pcscale.vendorkeys where counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend"))
            MyConnObj.Execute cm
            cm = "delete from pcscale.vendors where code=" & CStr(MyRecVar.Fields("ident_vend"))
            MyConnObj.Execute cm
            'vendorregister
            cm = "delete from pcscale.vendorregister where vendor=" & CStr(MyRecVar.Fields("ident_vend"))
            MyConnObj.Execute cm
            
            If MyRecVar.Fields("borrado") = True Then
                MyRecVar.Delete
            End If
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where tran_vend<>'*'")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from vendedor")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo CargaCadena(511) & " " & CStr(MyRecVar.Fields("ident_vend"))
'DROP TABLE IF EXISTS `pcscale`.`vendorkeys`;
'CREATE TABLE  `pcscale`.`vendorkeys` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Counter` int(11) DEFAULT NULL,
'  `VendorKey` int(11) DEFAULT NULL,
'  `Vendor` int(11) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Vendor` (`Vendor`),
'  UNIQUE KEY `Counter` (`Counter`,`VendorKey`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
            'MyRecDes.Open "Select * from pcscale.vendorkeys where vendor=" & CStr(MyRecVar.Fields("ident_vend")) & " and counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecVar.Fields("tec_vend") <> 0 Then
                MyRecDes.Open "Select * from pcscale.vendorkeys where counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    Call frmControl.Balanza_ComError(0)
                    GoTo FINVENDEDOR
                End If
                If MyRecDes.EOF Then
                    cm = "delete from pcscale.vendorkeys where vendor=" & CStr(MyRecVar.Fields("ident_vend"))
                    MyConnObj.Execute cm
                
                    MyRecDes.AddNew
                    lEsAlta = True
                Else
                    cm = "delete from pcscale.vendorkeys where counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend"))
                    MyConnObj.Execute cm

                    MyRecDes.AddNew
                    lEsAlta = True

                End If
    
                MyRecDes.Fields("vendor") = MyRecVar.Fields("ident_vend")
                MyRecDes.Fields("counter") = MyRecVar.Fields("secc_maqui")
                MyRecDes.Fields("vendorkey") = MyRecVar.Fields("tec_vend")
                MyRecDes.Update
                MyRecDes.Close
                Set MyRecDes = Nothing
                Do_Events
                
            End If
                
            On Error Resume Next
            MyRecDes.Open "select * from pcscale.vendors where code=" & CStr(MyRecVar.Fields("ident_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINVENDEDOR
            End If
            On Error GoTo FINVENDEDOR
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            MyRecDes.Fields("code") = MyRecVar.Fields("ident_vend")
            MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nombre")))

            MyRecDes.Update
            
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_vend") = "*"
            MyRecVar.Fields("prog_vend") = "*"
            
            MyRecVar.Update
            
            Do_Events
            cm = "delete from pcscale.vendorregister where vendor=" & CStr(MyRecVar.Fields("ident_vend"))
            MyConnObj.Execute cm
            
            MyRecVar.Movenext

            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirVend


        Loop
    
    
    End If
SalirVend:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINVENDEDOR:
    CadenadeLog "Error en MUE_EnviaVendedor:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
       
End Sub

Public Function Capa1A_GA_Erase_Totals(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para TouchScale  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim cm As String
    Dim nf As Integer
    Dim nC As Long
    Dim ncont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
    Dim nAct As Integer
    Dim spa As String
    Dim Arch As Integer
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0

    CadenadeLog "Capa1A_GA_Erase_totals"
'*******
'//Reinicio numeraciÃ³n
'SQL_SendMessage("update numberingtickets set resetnumber=1");
'    mysql_free_result(RespSQL); // Liberamos la variable pues no se necesita respuesta
'SQL_SendMessage("update erasetickets set dattim=now() where type='T'");
'    mysql_free_result(RespSQL); // Liberamos la variable pues no se necesita respuesta
'//////////////////////
'SQL_SendMessage("truncate numtickets");
'mysql_free_result(RespSQL);
'SQL_SendMessage("truncate localnumtickets");
'mysql_free_result(RespSQL);
        
    
    CadenadeLog "Iniciar Numeración si Procede..."
    If No_Reinit = 0 Then 'Iniciar numeración tiques...
        
        On Error Resume Next
        cm = "update pcscale.numberingtickets set resetnumber=1"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "update pcscale.erasetickets set dattim=now() where type='T'"
        MyConnObj.Execute cm
        On Error GoTo 0
        
        On Error Resume Next
        cm = "delete from pcscale.localnumtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "delete from pcscale.numtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        'On Error Resume Next
        'cm = "delete from pcscale.erasetickets"
        'MyConnObj.Execute cm
        'On Error GoTo 0
    End If
    CadenadeLog "Borrado Documentos y Lineas Documento..."
    On Error Resume Next
    If Dir(App.Path & "\nopark") = "" Then
        cm = "delete from pcscale.HTickets"
    Else
        cm = "delete from pcscale.HTickets where status<>'Z'"
    End If
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.sservicelabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.packetlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    'Añadididas 1.0.4
    cm = "delete from pcscale.localerasetickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localhtickets where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketpresets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpaytickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localsservicelabels"
    MyConnObj.Execute cm
    ''''''''''''''''''''''''''''''''
    On Error GoTo 0
    
    'BorrarInventarioFin 'BorrarInventario
    If BorrarInventarioFin Then  '1.9.1 j.mer.
         On Error Resume Next
        cm = "delete from pcscale.inventory"
        MyConnObj.Execute cm
        ''''''''''''''''''''''''''''''''
        On Error GoTo 0
        On Error Resume Next
        'inventory_weight, inventory_units
        cm = "update pcscale.stock set inventory_weight=0 where code<>0"
        MyConnObj.Execute cm
        ''''''''''''''''''''''''''''''''
        On Error GoTo 0
        On Error Resume Next
        'inventory_weight, inventory_units
        cm = "update pcscale.stock set inventory_units=0 where code<>0"
        MyConnObj.Execute cm
        ''''''''''''''''''''''''''''''''
        On Error GoTo 0
       
        
    End If
    If SC10_BorrarTotal Then
    'totaldaily
    'totalpaytickets
    'totalscalefamilies
    
         On Error Resume Next
        cm = "delete from pcscale.totaldaily"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
    
        cm = "delete from pcscale.totalpaytickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
    
        cm = "delete from pcscale.totalscalefamilies"
        MyConnObj.Execute cm
        On Error GoTo 0
       
    End If
    If SC10_Borrar_Eti_dat Then 'totalvat
    
        On Error Resume Next
        cm = "delete from pcscale.totalvat"
        MyConnObj.Execute cm
        On Error GoTo 0
        
    
    End If
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    If Resp = 0 Then '''ver tienda
        If TiendaActual = 0 Then
            spa = App.Path & "\"
        Else
            spa = App.Path & "\t" & Format(TiendaActual, "00") & "\"
        End If
        Arch = FreeFile()
        If Dir(spa & "ngtord") = "" Then
            Open spa & "ngtord" For Output As #Arch
            Print #Arch, "000001"
            nGTOrd = 1
        Else
            Open spa & "ngtord" For Input As #Arch
            Line Input #Arch, cm
            Close #Arch
            If Val(cm) > 9999 Then cm = "00000"
            Open spa & "ngtord" For Output As #Arch
            Print #Arch, Format(Val(cm) + 1, "000000")
            nGTOrd = Val(cm) + 1
        End If
        Close #Arch
    End If
    
    Capa1A_GA_Erase_Totals = Resp
End Function

''''¡¡¡¡¡¡¡¡¡¡¡¡¡ Solo utilizar en final día!!!!!!!!!!!!!!!!!!!'''''''
'   Se borran datos de la Balanza
'   Se deben introducir tiques en dbasetouch.mdb con typtic=4
'   En principio la información de cabecera no es válida...
'   Se pensaba en utilizar las tablas de TouchScale.MDB pero se va a traspasar directamente a dbasetouch.mdb
Public Function Tiques(ByVal lLocal As Boolean) As Integer
Dim MyBaseORG As dao.Database
Dim MyRsORG As dao.Recordset

Dim MyBaseORG1 As dao.Database
Dim MyRsORG1 As dao.Recordset
Dim MyrsART As dao.Recordset
Dim MyRsVEND As dao.Recordset

Dim MyBaseBAL As dao.Database
Dim MyRsBAL As dao.Recordset
Dim siP As String
Dim MyRsDes As New ADODB.Recordset
Dim myRsKEYS As New ADODB.Recordset
Dim myRsTMP As New ADODB.Recordset
Dim cm As String
Dim nf As Integer
Dim ncont As Integer
Dim sS As String
Dim nImporte As Long
Dim sHour As String
Dim sfecha As String
'Dim nF As Integer

'    lComunica = True
'nF = FreeFile()
'Open App.Path & "\LOGTOUCH.TXT" For Output As #nF
'Print #nF, "Recogida Tiques Touch-Scale..."
    
    sHour = Format(Time, "hh:mm:ss")
    sfecha = Format(Date, "dd/mm/yy")
    
    'Set MyBaseBAL = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    'Set MyRsBAL = MyBaseBAL.OpenRecordset("select * from balanzas where ip<>'127-0-0-1' order by scalenr")
    Set MyBaseORG = dao.OpenDatabase(App.Path & "\dbasetouch.mdb")
    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where borrado=false")
    ''Call AbrirTotales
    ''If lTotalesAbierta Then
    ''    nF = nF
    ''End If
    ''Cm = "select * from articulo where borrado=false"
    ''MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic


    'If Not MyRsBAL.EOF Then
        'MyRsBAL.MoveFirst
        'Do While Not MyRsBAL.EOF
        
            'On Error GoTo ProblemaPing
        
            'sIp = toIpAdd(Trim(MyRsBAL.Fields("ip")))
            
            
            'If ElPing(sIp) = 1 Then
    
            '    Call AbrirMySQL(sIp)
                
            '    If lMySqlAbierta Then
                    For ncont = 0 To 1 '0=local, 1=Red
                        If ncont = 0 Then
'Print #nF, "Recogida Tiques local Touch-Scale..."
'DROP TABLE IF EXISTS `pcscale`.`htickets`;
'CREATE TABLE  `pcscale`.`htickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdNetHticket` int(11) DEFAULT '0',
'  `NetStat` char(1) DEFAULT '',
'  `Captured` char(1) DEFAULT '',
'  `tckType` char(1) DEFAULT NULL,
'  `Status` char(1) DEFAULT NULL,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `DatTim` datetime DEFAULT NULL,
'  `ReloadFrom` int(11) DEFAULT '0',
'  `Customer` int(11) DEFAULT '0',
'  `tckDiscount` decimal(4,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Vendor` (`Vendor`,`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                            
'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
'CREATE TABLE  `pcscale`.`ltickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                            cm = "select * from pcscale.LocalLTickets"
                            MyRsDes.Open cm, MyConnObj, adOpenDynamic, adLockOptimistic
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF

                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG!seccion = CStr(8)
                                    MyRsORG!Balanza = MyRsBAL!scalenr
                                    MyRsORG!vended = CStr(6099)
                                    MyRsORG!numlin = CStr(1)
                                    MyRsORG!lincan = CStr(0)
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG!Importe = MyRsDes!Price
                                        nImporte = (MyRsDes!Price) * 100
                                    Else
                                        MyRsORG!Importe = Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
                                        nImporte = 100 * (Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2))
                                    End If
                                    MyRsORG!hora = sHour
                                    MyRsORG!onoff = "Local"
                                    MyRsORG!STPETIC = CStr(16)
                                    MyRsORG!codcli = CStr(0)
                                    MyRsORG!Fecha = sfecha
                                    MyRsORG!TYPTIC = CStr(4)
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
                                    sS = sS & CStr(DigitoControl(sS))
                                    MyRsORG!codbar = sS
                                    MyRsORG!reserv1 = "No Enviado"
                                    MyRsORG!reserv2 = "*"
                                    MyRsORG!reserv3 = "SI"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sfecha
                                    MyRsORG1!Hour = sHour
                                    MyRsORG1!numbal = MyRsBAL!scalenr
                                    MyRsORG1!Vendor = CStr(6099)
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG1!numlin = CStr(1)
                                    MyRsORG1!secc = CStr(8)
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    
                                    If MyrsART.EOF Then
                                        MyRsORG1!Plu = CStr(0)
                                        MyRsORG1!Item = ""
                                    Else
                                        MyRsORG1!Plu = CStr(MyrsART!Plu)
                                        MyRsORG1!Item = MyrsART!des_plu1
                                    End If
                                    MyRsORG1!Price = MyRsDes!Price
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG1!amount = MyRsDes!Price
                                        MyRsORG1!weigth = 0
                                        MyRsORG1!units = 1
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
                                        MyRsORG1.Fields("Weight") = MyRsDes.Fields("Weight")
                                        MyRsORG1!units = 0
                                    End If
                                    MyRsORG1!STPETIC = CStr(16)
                                    MyRsORG1!lincan = CStr(0)
                                    MyRsORG1!IVA = 0
                                    MyRsORG1!TYPTIC = CStr(4)
                                    
                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sfecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sfecha, "dd/mm/yy")
                                            !codi_ident = 188
                                            !ident_vend = 6099
                                            !secc_maqui = 8
                                            !Modo = 4
                                            !tiquets = 0
                                            !operacion = 0
                                            !cancela = 0
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !Peso = 0
                                            !unidad = 0
                                            ' ****
                                            !pesetas = 0
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
                                            !Supertot = 0
                                            !Total = 0
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sfecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
                                            !operacion = !operacion + 1
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!Price)
                                                !unidad = !unidad + 1
                                                !Total = !Total + (MyRsDes!Price)
                                                !Supertot = !Total + MyRsDes!Price
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
                                                !Peso = !Peso + (MyRsDes.Fields("Weight"))
                                                !Total = !Total + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
                                                !Supertot = !Total + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    
                                    MyRsDes.Movenext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                        
                            MyRsVEND.Clone
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        Else
                            
'Print #nF, "Recogida Tiques Red Touch-Scale..."
                            
                            cm = "select * from pcscale.LTickets"
                            MyRsDes.Open cm, MyConnObj, adOpenDynamic, adLockOptimistic
'Print #nF, "Paso 001"
                            
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
'Print #nF, "Paso 002"
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF
                                    
                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
'Print #nF, "Paso 003"
                                    MyRsORG!seccion = CStr(8)
'Print #nF, "Paso 004"
                                    MyRsORG!Balanza = 1 'Ver número Balanza
'Print #nF, "Paso 005"
                                    MyRsORG!vended = CStr(6099)
'Print #nF, "Paso 006"
                                    MyRsORG!numlin = CStr(1)
'Print #nF, "Paso 007"
                                    MyRsORG!lincan = CStr(0)
'Print #nF, "Paso 008"
                                    If MyRsDes!saleform = 0 Then
'Print #nF, "Paso 009"
                                        
                                        MyRsORG!Importe = MyRsDes!Price
'Print #nF, "Paso 0010"
                                        nImporte = (MyRsDes!Price) * 100
'Print #nF, "Paso 0011"
                                    Else
                                        MyRsORG!Importe = Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0012"
                                        nImporte = 100 * (Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2))
'Print #nF, "Paso 0013"
                                    End If
                                    MyRsORG!hora = sHour
'Print #nF, "Paso 0014"
                                    MyRsORG!onoff = "Red"
'Print #nF, "Paso 0015"
                                    MyRsORG!STPETIC = CStr(16)
'Print #nF, "Paso 0016"
                                    MyRsORG!codcli = CStr(0)
'Print #nF, "Paso 0017"
                                    MyRsORG!Fecha = sfecha
'Print #nF, "Paso 0018"
                                    MyRsORG!TYPTIC = CStr(4)
'Print #nF, "Paso 0019"
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
'Print #nF, "Paso 0020"
                                    sS = sS & CStr(DigitoControl(sS))
'Print #nF, "Paso 0021"
                                    MyRsORG!codbar = sS
'Print #nF, "Paso 0022"
                                    MyRsORG!reserv1 = "No Enviado"
'Print #nF, "Paso 0023"
                                    MyRsORG!reserv2 = "*"
'Print #nF, "Paso 0024"
                                    MyRsORG!reserv3 = "SI"
'Print #nF, "Paso 0025"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sfecha
'Print #nF, "Paso 0026"
                                    MyRsORG1!Hour = sHour
'Print #nF, "Paso 0027"
                                    MyRsORG1!numbal = 1 'MyRsBAL!scalenr
'Print #nF, "Paso 0028"
                                    MyRsORG1!Vendor = CStr(6099)
'Print #nF, "Paso 0029"
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
'Print #nF, "Paso 0030"
                                    MyRsORG1!numlin = CStr(1)
'Print #nF, "Paso 0031"
                                    MyRsORG1!secc = CStr(8)
'Print #nF, "Paso 0032"
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
'Print #nF, "Paso 0033"
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
'Print #nF, "Paso 0034"
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    If MyrsART.EOF Then
                                        MyRsORG1!Plu = CStr(0)
'Print #nF, "Paso 0035"
                                        MyRsORG1!Item = ""
'Print #nF, "Paso 0036"
                                    Else
                                        MyRsORG1!Plu = CStr(MyrsART!Plu)
'Print #nF, "Paso 0037"
                                        MyRsORG1!Item = MyrsART!des_plu1
'Print #nF, "Paso 0038"
                                    End If
                                    MyRsORG1!Price = MyRsDes!Price
'Print #nF, "Paso 0039"
                                    If MyRsDes!saleform = 0 Then
'Print #nF, "Paso 0040"
                                        MyRsORG1!amount = MyRsDes!Price
'Print #nF, "Paso 0041"
                                        MyRsORG1.Fields("weight") = 0
'Print #nF, "Paso 0042"
                                        MyRsORG1!units = 1
'Print #nF, "Paso 0043"
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0044"
                                        MyRsORG1.Fields("Weight") = MyRsDes.Fields("Weight")
'Print #nF, "Paso 0045"
                                        MyRsORG1!units = 0
'Print #nF, "Paso 0046"
                                    End If

                                    MyRsORG1!STPETIC = CStr(16)
'Print #nF, "Paso 0046"
                                    MyRsORG1!lincan = CStr(0)
'Print #nF, "Paso 0047"
                                    MyRsORG1!IVA = 0
'Print #nF, "Paso 0048"
                                    MyRsORG1!TYPTIC = CStr(4)
'Print #nF, "Paso 0049"
                                    

                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sfecha, "mm/dd/yy") & "#")
'Print #nF, "Paso 0050"
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sfecha, "dd/mm/yy")
'Print #nF, "Paso 0051"
                                            !codi_ident = 188
'Print #nF, "Paso 0052"
                                            !ident_vend = 6099
'Print #nF, "Paso 0053"
                                            !secc_maqui = 8
'Print #nF, "Paso 0054"
                                            !Modo = 4
'Print #nF, "Paso 0055"
                                            !tiquets = 0
'Print #nF, "Paso 0056"
                                            !operacion = 0
'Print #nF, "Paso 0057"
                                            !cancela = 0
'Print #nF, "Paso 0058"
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
'Print #nF, "Paso 0059"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !Peso = 0
'Print #nF, "Paso 0060"
                                            !unidad = 0
'Print #nF, "Paso 0061"
                                            ' ****
                                            !pesetas = 0
'Print #nF, "Paso 0062"
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
'Print #nF, "Paso 0063"
                                            !Supertot = 0
'Print #nF, "Paso 0064"
                                            !Total = 0
'Print #nF, "Paso 0065"
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sfecha, "mm/dd/yy") & "#")
'Print #nF, "Paso 0066"
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
'Print #nF, "Paso 0067"
                                            !operacion = !operacion + 1
'Print #nF, "Paso 0068"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
'Print #nF, "Paso 0069"
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!Price)
'Print #nF, "Paso 0070"
                                                !unidad = !unidad + 1
'Print #nF, "Paso 0071"
                                                !Total = !Total + (MyRsDes!Price)
'Print #nF, "Paso 0072"
                                                !Supertot = !Total + MyRsDes!Price
'Print #nF, "Paso 0073"
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0074"
                                                !Peso = !Peso + (MyRsDes.Fields("Weight"))
'Print #nF, "Paso 0075"
                                                !Total = !Total + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0076"
                                                !Supertot = !Total + Round((MyRsDes!Price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0077"
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    MyRsDes.Movenext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                            MyRsVEND.Close
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        End If
                    
                    Next ncont
                            
                    'Borrado de los datos de Totales en Balanza
'Print #nF, "Borrado Tiques Touch-Scale..."
                    On Error Resume Next
                    cm = "delete from pcscale.paytickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.HTickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.LTickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.sservicelabels"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.packetslabels"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    
                    'Cm = "delete from pcscale.LocalHTickets"
                    'myconnobj.Execute Cm
                    'Cm = "delete from pcscale.LocalLTickets"
                    'myconnobj.Execute Cm
                    'Cm = "delete from pcscale.localnumtickets"
                    'myconnobj.Execute Cm
                    
                    'Intento reiniciar numeración
                    'Cm = "update pcscale.status set resetticketnumber=1 where id=1"
                    'myconnobj.Execute Cm
                    'Cm = "update pcscale.erasetickets set dattim=now()"
                    'myconnobj.Execute Cm
                    On Error Resume Next
                    cm = "delete from pcscale.localnumtickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.numtickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    On Error Resume Next
                    cm = "delete from pcscale.erasetickets"
                    MyConnObj.Execute cm
                    On Error GoTo 0
                    
                    ''''''''''''''''''''''''''''''

            '    End If
            
            'End If
        
            'If lMySqlAbierta Then
            '    myconnobj.Close
            '    Set DbMysqli = Nothing
            '    lMySqlAbierta = False
            'End If
            '
            'MyRsBAL.Movenext
        
        'Loop
        
    'End If

    On Error Resume Next
    
    'MyRsBAL.Close
    'Set MyRsBAL = Nothing
    
    MyBaseBAL.Close
    Set MyBaseBAL = Nothing
    
    MyrsART.Close
    Set MyrsART = Nothing

    MyBaseORG.Close
    Set MyBaseORG = Nothing
    'If lTotalesAbierta Then
    '    DbTotales.Close
    '    Set DbTotales = Nothing
    'End If
    
'    On Error GoTo 0
'ProblemaPing:
    'If Err.Number <> 0 Then
        'Print #nF, "Error:" & CStr(Err.Number) & " Desc.:" & Err.Description
        'On Error GoTo 0
        'On Error Resume Next
        'If lMySqlAbierta Then
        '    myconnobj.Close
        '    Set DbMysqli = Nothing
        'End If
        'On Error GoTo 0
        
    'End If
    'On Error Resume Next
    'Close #nF
    On Error GoTo 0
    'lComunica = False
    
End Function


'DROP TABLE IF EXISTS `pcscale`.`htickets`;
'CREATE TABLE  `pcscale`.`htickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdNetHticket` int(11) DEFAULT '0',
'  `NetStat` char(1) DEFAULT '',
'  `Captured` char(1) DEFAULT '',
'  `CapturedDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `tckType` char(1) DEFAULT NULL,
'  `Status` char(1) DEFAULT NULL,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `NumLines` int(11) DEFAULT '0',
'  `AddedLines` int(11) DEFAULT '0',
'  `GrossAmount` decimal(8,2) DEFAULT '0.00',
'  `NetAmount` decimal(8,2) DEFAULT '0.00',
'  `EAN13` char(13) NOT NULL DEFAULT '',
'  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `TimesPrinted` tinyint(3) unsigned NOT NULL DEFAULT '0',
'  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `ReloadFrom` int(11) DEFAULT '0',
'  `ReloadScale` int(11) NOT NULL DEFAULT '0',
'  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
'  `Customer` varchar(16) NOT NULL DEFAULT '',
'  `Invoice` varchar(20) NOT NULL DEFAULT '',
'  `tckDiscount` decimal(4,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  KEY `Vendor` (`Vendor`),
'  KEY `Invoice` (`Invoice`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
'CREATE TABLE  `pcscale`.`ltickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `Name` varchar(127) NOT NULL DEFAULT '',
'  `Section` decimal(3,0) NOT NULL DEFAULT '0',
'  `LineScale` int(11) NOT NULL DEFAULT '0',
'  `LineVendor` int(11) NOT NULL DEFAULT '0',
'  `LineDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `SourceKey` char(1) NOT NULL DEFAULT '0',
'  `PriceChanged` tinyint(1) NOT NULL DEFAULT '0',
'  `VATCode` decimal(1,0) unsigned NOT NULL DEFAULT '0',
'  `VATPercent` decimal(4,2) unsigned NOT NULL DEFAULT '0.00',
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  `FreezingDate` date DEFAULT NULL,
'  `TicketDiscount` decimal(4,2) DEFAULT '0.00',
'  `PrintScale` int(11) NOT NULL DEFAULT '0',
'  `PrintDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `PrintTicketNumber` int(11) NOT NULL DEFAULT '0',
'  `ReloadScale` int(11) NOT NULL DEFAULT '0',
'  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`),
'  CONSTRAINT `ltickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `htickets` (`Id`) ON DELETE CASCADE
') ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;

'CREATE TABLE  `pcscale`.`paytickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `PayMethod` int(11) DEFAULT NULL,
'  `Delivered` decimal(8,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`paymethods`;
'CREATE TABLE  `pcscale`.`paymethods` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Name` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Name` (`Name`)
') ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`packetlabels`;
'CREATE TABLE  `pcscale`.`packetlabels` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `Price` decimal(6,2) DEFAULT NULL,
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  `FreezingDate` date DEFAULT NULL,
'  `Customer` varchar(16) NOT NULL DEFAULT '',
'  `DatTim` datetime DEFAULT NULL,
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`sservicelabels`;
'CREATE TABLE  `pcscale`.`sservicelabels` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `Price` decimal(6,2) DEFAULT NULL,
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  `FreezingDate` date DEFAULT NULL,
'  `Customer` varchar(16) NOT NULL DEFAULT '',
'  `DatTim` datetime DEFAULT NULL,
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;

'Cabeceras de tique:
'.- NetStat: ‘N’, ‘L’.
'    Tique en red (‘N’) o local (‘L’).
'    Los tiques de red copiados en local se guardan en la tabla LocalHTicktes con NetStat=’N’.
'.- TckType: ‘N’, ‘O’
'    Tique “Normal” o  “Order” (encargo).
'    Los tiques encargo sólo se marcan como tales, no se realiza ninguna gestión especial.
'.- Status: ‘O’, ‘P’, ‘C’, ’D’, ‘R’, ‘N’, ‘X’, 'Z'
'    Estado del tique: abierto (‘O’), en impresión (‘P’), cerrado (‘C’), cancelado (‘D’),
'    recargado en tique abierto (‘R’),  recargado en tique cerrado (‘N’), cancelado manualmente (‘X’)
'    ('Z') Aparcado
'Líneas de tique:
'.- LineType: ‘1’, ‘0’
'    Línea “normal” (‘1’) o cancelada (‘0’).
'.- TTare: ‘ ‘ , ‘W’, ‘M’, ‘I’
'    Tipo de tara aplicado en la operación. Tara acumulativa/pesada (‘W‘), manual (‘M’),
'    programada en el artículo (‘I’), sin tara (‘ ‘)


Public Sub LeeTickets(Optional Tabla As Integer, Optional ByVal lFinDia As Boolean)
    
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    Dim MyRecImp As New ADODB.Recordset
    
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Dim MisEqp As dao.Recordset
    Dim nLR As Integer
    
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim ncont As Long
    Dim dLImporte As Double
    Dim nLocal As Integer
    
    'If Dir(App.Path & "\hydramulti.exe") <> "" Or Dir(App.Path & "\noslaves.txt") <> "" Then
    '    nLocal = 0
    'Else
        nLocal = 1
    'End If
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'Call capturaClientes
    
    On Error GoTo fin
    Set Mybase = dao.OpenDatabase(Base_General)
    'CadenadeLog "paso 1"
    'If nLocal <> 0 Then
    If Dir(App.Path & "\tqslaves.txt") <> "" Then
        Set MisEqp = Mybase.OpenRecordset("select * from equipos order by modelo asc")
    Else
        Set MisEqp = Mybase.OpenRecordset("select * from equipos where modelo=1")
    End If
    'CadenadeLog "paso 2"
    If MisEqp.EOF Then
        GoTo SalirLeeTique
    Else
        MisEqp.MoveFirst
        Do While Not MisEqp.EOF
            
            'If Dir(App.Path & "\tqslaves.txt") <> "" Then
            '    frmControl.Switch_M MisEqp.Fields("prog_eqp")
            'Else
            '    frmControl.Switch_mode 0
            'End If
            If MisEqp.Fields("modelo") = 1 Then
                frmControl.Switch_mode 0
            Else
                frmControl.Switch_M MisEqp.Fields("prog_eqp")
            End If
            'CadenadeLog "paso 3"
            If EstadoBal = 0 Then
            
            For nLR = 0 To 1 ' Red/Local
            
            If MisEqp.Fields("modelo") <> 1 And nLR = 0 Then GoTo SigueNlr
            'CadenadeLog "paso 4"
            If nLR = 0 Then
                If IsNull(Tabla) Or Tabla = 0 Then
                    MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or captured<>'*') and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    'CadenadeLog "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='N' order by id"
                Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.packetlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.packetlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.sservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.sservicelabels where customer<>'9999999999' order by id"
                    End If
                End If
            Else
                If IsNull(Tabla) Or Tabla = 0 Then
                    'MyRecSetTmp.Open "select * from pcscale.localhtickets where captured='' and status='C' and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or captured<>'*') and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    'CadenadeLog "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='L' order by id"
                Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id"
                    End If
                End If
            End If
            'CadenadeLog "paso 5"
            If Not MyRecSetTmp.EOF Then
                Do While Not MyRecSetTmp.EOF
                    'CadenadeLog "paso 6"
                    Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
                    Fhora = Mid(Ffecha, 12, 8)
                    Ffecha = Mid(Ffecha, 1, 10)
                    'CadenadeLog "paso 7"
                    IdDoc = MyRecSetTmp.Fields("id")
                    'CadenadeLog "paso 8"
                    ''''medio de pago
                    '1=efectivo
                    '2=t.crédito
                    '3=cheque
                    Cabecera.tipoPago = 0 'era 1 (pasa a 0 efectivo) 3.4.0
                    Cabecera.cantidadEntre = 0
                    
                    If IsNull(Tabla) Or Tabla = 0 Then
'Select Case !TYPTIC --> Cabecera.tipoTiquet
'    Case 1
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
'    Case 2
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
'    Case 3
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
'    Case 4
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
'    Case 5
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
'    Case 6
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
'    Case 7
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
'    Case 8
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
'    Case 12
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Descart."
'End Select
                    
                        'If nLR = 0 Then
                            MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc)
                            If Not MyRecAux.EOF Then
                                'CadenadeLog "paso 9"
                                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                                'CadenadeLog "paso 10"
                            End If
                            MyRecAux.Close
                            Set MyRecAux = Nothing
'.- TckType: ‘N’, ‘O’
'    Tique “Normal” o  “Order” (encargo).
                            If MyRecSetTmp.Fields("tcktype") = "O" Then Cabecera.tipoPago = 12 '3.4.0 encargo (Order)
                            
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'CadenadeLog "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc)
                        '    If Not MyRecAux.EOF Then
                        '        'CadenadeLog "paso 11"
                        '        Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                        '        Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                        '        'CadenadeLog "paso 12"
                        '    End If
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                    End If
        '''''''
                    'Impuestos
                    Cabecera.base0 = 0
                    Cabecera.por0 = 0
                    Cabecera.imp0 = 0
                    Cabecera.base1 = 0
                    Cabecera.por1 = 0
                    Cabecera.imp1 = 0
                    Cabecera.base2 = 0
                    Cabecera.por2 = 0
                    Cabecera.imp2 = 0
                    Cabecera.base3 = 0
                    Cabecera.por3 = 0
                    Cabecera.imp3 = 0
                    Cabecera.base4 = 0
                    Cabecera.por4 = 0
                    Cabecera.imp4 = 0
        ''''''''''
                    
                    '''''''''''''''''
                    If IsNull(Tabla) Or Tabla = 0 Then
                        
                        'If nLR = 0 Then
                            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "paso 13"
                            nNumLin = MyRecAux.Fields("lineas")
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            ReDim Cabecera.Lines(nNumLin)
                            'MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                            'On Error Resume Next
                            ''CadenadeLog "paso 14"
                            'dLImporte = MyRecAux.Fields("importe")
                            ''CadenadeLog "paso 15"
                            'If Err.Number <> 0 Then
                            '    On Error GoTo 0
                            '    On Error GoTo FIN
                            '    dLImporte = 0
                            'End If
                            'MyRecAux.Close
                            'Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    'CadenadeLog "paso 16"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '    ReDim Cabecera.Lines(nNumLin)
                        '    'MyRecAux.Open "select sum(amount) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'On Error Resume Next
                        '    'dLImporte = MyRecAux.Fields("importe")
                        '    ''CadenadeLog "paso 17"
                        '    'If Err.Number <> 0 Then
                        '    '    On Error GoTo 0
                        '    '    On Error GoTo FIN
                        '    '    dLImporte = 0
                        '    'End If
                        '    'MyRecAux.Close
                        '    'Set MyRecAux = Nothing
                        '
                        'End If
                    Else
                        nNumLin = 1
                        'CadenadeLog "paso 18"
                        ReDim Cabecera.Lines(nNumLin)
                        'dLImporte = MyRecSetTmp.Fields("amount")
                        'CadenadeLog "paso 19"
                    End If
                    
                    'recoger datos
                    Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
                    Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
                    Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
                    'CadenadeLog "paso 20"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        If IsNull(MyRecSetTmp.Fields("invoice")) Or Len(MyRecSetTmp.Fields("invoice")) = 0 Then
                            Cabecera.factura = Space(15)
                        Else
                            Cabecera.factura = Trim(MyRecSetTmp.Fields("invoice"))
                            'CadenadeLog "paso 21"
                        End If
                        ''''''''''
                        Cabecera.TotalUnits = MyRecSetTmp.Fields("totalunits")
                        Cabecera.TotalWeight = MyRecSetTmp.Fields("totalweight")
                        ''''''''''
                        Cabecera.ean13 = Trim(MyRecSetTmp.Fields("ean13"))
                        'CadenadeLog "paso 22"
                        If MyRecSetTmp.Fields("customer") <> "" Then
                            If IsNull(MyRecSetTmp.Fields("customer")) Then
                                Cabecera.NCliente = 0
                            Else
                                On Error Resume Next
                                Cabecera.NCliente = MyRecSetTmp.Fields("customer")
                                If Err.Number <> 0 Then
                                    On Error GoTo 0
                                    Cabecera.NCliente = 0
                                End If
                                On Error GoTo fin
                            End If
                        Else
                            Cabecera.NCliente = 0
                        End If
                        'CadenadeLog "paso 23"
                    Else
                        Cabecera.factura = Space(15)
                        Cabecera.NCliente = 0
                    End If
                    'CadenadeLog "paso 24"
                    
                    Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
                    Cabecera.hora = Fhora
                    Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
                    If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("netamount")
                    Else
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("amount")
                    End If
                    'CadenadeLog "paso 25"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.nLineas = nNumLin
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = nNumLin
                        Cabecera.LineasMensaje = nNumLin
                    Else
                        Cabecera.nLineas = 1
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = 1
                        Cabecera.LineasMensaje = 1
                    End If
                    'cabecera.codigo ???
                    
                    Cabecera.EstadoTiquet = 16
                    'CadenadeLog "paso 26"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If MyRecSetTmp.Fields("NetStat") = "L" Then
                            Cabecera.CadenaEstadoTiquet = "Local"
                            Cabecera.tipoTiquet = 4
                        Else
                            Cabecera.CadenaEstadoTiquet = "Red"
                            Cabecera.tipoTiquet = 1
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'CadenadeLog "paso 27"
                        'If nLR = 0 Then
                        '    'CadenadeLog "paso 28"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 29"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    'CadenadeLog "paso 30"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 31"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                        'CadenadeLog "paso 32"
''V': Tique reembolso (cantidades vendidas negativas). (SAFT_PT) (+DB.135)
''w: Tique reembolsado (cancelado) pero no recargado. (SAFT_PT) (+DB.135)
''W': Tique reembolsado (cancelado) y recargado. (SAFT_PT) (+DB.135) .

''X': Cancelado por el operador. Sin imprimir.
''R': Tique recargado, Tique “Cerrado” copiado sobre un nuevo tique para añadir nuevas operaciones.

                        Cabecera.Trainning = False
                        If MyRecSetTmp.Fields("tcktype") = "U" Then
                            Cabecera.tipoTiquet = 12
                            Cabecera.Trainning = True
                        End If
                        If MyRecSetTmp.Fields("status") = "R" Or MyRecSetTmp.Fields("status") = "V" Then
                            Cabecera.tipoTiquet = 13
                            Cabecera.EstadoTiquet = 64

                        End If
                        'CadenadeLog "paso 33"
                        If MyRecSetTmp.Fields("status") = "X" Or MyRecSetTmp.Fields("status") = "W" Or MyRecSetTmp.Fields("status") = "w" Then
                            Cabecera.tipoTiquet = 14
                            Cabecera.EstadoTiquet = 64
                            
                        End If
                        'CadenadeLog "paso 34"
                        Cabecera.ImporteBruto = MyRecSetTmp.Fields("grossamount")
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Round(Cabecera.ImporteBruto - Cabecera.ImporteTotal, 2) 'más abajo se hace como porcentaje...
                        Cabecera.documentoabono = 0
                    
                    Else
                        'CadenadeLog "paso 35"
                        If nLR = 0 Then
                            Cabecera.CadenaEstadoTiquet = "Red"
                        Else
                            Cabecera.CadenaEstadoTiquet = "Local"
                        End If
                        
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If Tabla = 1 Then
                            Cabecera.tipoTiquet = 2
                        Else
                            Cabecera.tipoTiquet = 3
                        End If
                        
                        'Cabecera.CadenaEstadoTiquet = "Red"
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        'Cabecera.ImporteBruto = dLImporte
                        Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal
                        Cabecera.documentoabono = 0
                        'Cabecera.Cambio = 0
                        'Cabecera.descuento = 0
                        'Cabecera.LineasCanceladas = 0
                    End If
                    'CadenadeLog "paso 36"
                    If nLR = 0 Then
                        If Tabla = 0 Or IsNull(Tabla) Then
                            frmMonitor.MostrarDato CargaCadena(868) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            frmMonitor.MostrarDato CargaCadena(1412) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            frmMonitor.MostrarDato CargaCadena(1413) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    Else
                        If Tabla = 0 Or IsNull(Tabla) Then
                            frmMonitor.MostrarDato CargaCadena(1414) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            frmMonitor.MostrarDato CargaCadena(1415) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            frmMonitor.MostrarDato CargaCadena(1416) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    
                    End If
                    'CadenadeLog "paso 37"
                    If Tabla = 0 Or IsNull(Tabla) Then
                        'If nLR = 0 Then
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '
                        'End If
                        'CadenadeLog "paso 38"
                        'CadenadeLog "7"
                        
                        Cabecera.LineasCanceladas = MyRecSetTmp.Fields("canceledlines")
                        
                        'también se vé afectado si es cancelado o no pero de mamento cerrado...
                        'Cabecera.EstadoTiquet = 16
                        '********************
                        ' este dato se introduce
                        ' tras pasar por actualiza_tabla_cabecera
                        'Incluir_en_Base As Boolean
                        'Cabecera.tipoPago = myrecsettmp.Fields("idmoneda")
                        'Cabecera.cantidadEntre = myrecsettmp.Fields("totalneto") + myrecsettmp.Fields("importecambio")
                        If Cabecera.cantidadEntre > 0 Then
                            Cabecera.Cambio = Round(Cabecera.ImporteTotal - Cabecera.cantidadEntre, 2) ' myrecsettmp.Fields("importecambio")
                        Else
                            Cabecera.Cambio = 0
                        End If
                        Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")
                        'CadenadeLog "paso 39"
                        'lote As String
                        'lEncargo As Boolean
                                
                        If lFornes Then
                            If MyRecSetTmp.Fields("usesectionitem") = 1 Then
                                Cabecera.NMostrador = MyRecSetTmp.Fields("sectionitem")
                            End If
                        End If
            
            
                        'LINEAS
            
                        nNumLin = 0
                        'If nLR = 0 Then
                            MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'End If
                        'CadenadeLog "8"
                        'CadenadeLog "paso 40"
                        If Not MyRecAux.EOF Then
                            
                            'CadenadeLog "8.1"
                            
                            MyRecAux.MoveFirst
                            
                            'CadenadeLog "8.2"
                            
                            Do While Not MyRecAux.EOF
                                
                                'CadenadeLog "8.3"
            'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
            'CREATE TABLE  `pcscale`.`ltickets` (
            '  `Id` int(11) NOT NULL AUTO_INCREMENT,
            '  `IdHTicket` int(11) DEFAULT NULL,
            '  `lineType` smallint(6) DEFAULT NULL,
            '  `SaleForm` smallint(6) DEFAULT NULL,
            '  `Item` int(11) DEFAULT NULL,
            '  `Weight` decimal(5,3) DEFAULT NULL,
            '  `Tare` decimal(5,3) DEFAULT '0.000',
            '  `TTare` char(1) DEFAULT '',
            '  `lineDiscount` decimal(4,2) DEFAULT NULL,
            '  `Price` decimal(6,2) DEFAULT NULL,
            '  `Amount` decimal(8,2) DEFAULT NULL,
            '  `BovineCard` smallint(3) DEFAULT NULL,
            '  `EarTag` varchar(30) DEFAULT NULL,
            '  `BovineLot` varchar(20) DEFAULT NULL,
            '  PRIMARY KEY (`Id`),
            '  KEY `IdHTicket` (`IdHTicket`)
            ') ENGINE=InnoDB DEFAULT CHARSET=latin1;
                                'CadenadeLog "paso 41"
                                nNumLin = nNumLin + 1
                                If MyRecAux.Fields("saleform") = 1 Then
                                    Cabecera.Lines(nNumLin).EsPesado = True
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                Else
                                    Cabecera.Lines(nNumLin).EsPesado = False
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                End If
                                'CadenadeLog "paso 42"
                                'CadenadeLog "8.4"
                                
                                Cabecera.Lines(nNumLin).FactorPeso = 1
                                Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                                'CadenadeLog "paso 43"
                                'CadenadeLog "8.5"
                                
                                Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                                Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                                '3.4.4
                                Cabecera.Lines(nNumLin).desc = MyRecAux.Fields("name")
                                ''''''
                                'CadenadeLog "paso 44"
                                'ver qué se hace IVA As Integer
                                Cabecera.Lines(nNumLin).IVA = MyRecAux.Fields("vatcode")
                                Cabecera.Lines(nNumLin).porcentaje = MyRecAux.Fields("vatpercent")
                                'CadenadeLog "paso 45"
                                'CadenadeLog "8.6"
                                
                                Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("linediscount") ' MyRecAux.Fields("ticketdiscount") + MyRecAux.Fields("linediscount") 'porcentaje debido a la cabecera + porcentaje en linea
                                'CadenadeLog "paso 46"
                                Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2) 'MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                                'CadenadeLog "paso 47"
                                Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("amount") '- Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                
                                If Cabecera.Lines(nNumLin).ImporteBruto > Cabecera.Lines(nNumLin).ImporteLinea Then
                                
                                    Cabecera.ImporteBruto = Cabecera.ImporteBruto + (Cabecera.Lines(nNumLin).ImporteBruto - Cabecera.Lines(nNumLin).ImporteLinea)
                                
                                End If
                                
                                
                                'Cabecera.Lines(nNumLin).ImporteLinea = Round(Cabecera.Lines(nNumLin).ImporteLinea - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                
                                'CadenadeLog "paso 48"
                                'CadenadeLog "8.7"
                                
                                'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                                If MyRecAux.Fields("linetype") = 0 Then
                                    Cabecera.Lines(nNumLin).LineaCancelada = True
                                Else
                                    Cabecera.Lines(nNumLin).LineaCancelada = False
                                End If
                                'CadenadeLog "paso 49"
                                'ver qué se hace con TipoLinea As Byte
                                'STV (positivo o negativo)
                                'Linea(Bucle).ModificadoTemporal = False
                                If MyRecAux.Fields("amount") > 0 Then
                                    'Cabecera.Lines(nNumLin).positive = 1
                                     Cabecera.Lines(nNumLin).FactorPeso = 1
                                     Cabecera.Lines(nNumLin).TipoLinea = 0
                                
                                Else
                                    'Cabecera.Lines(nNumLin).positive = 0
                                     Cabecera.Lines(nNumLin).FactorPeso = -1
                                     Cabecera.Lines(nNumLin).TipoLinea = 2
                                    
                                End If
                                'CadenadeLog "paso 50"
            '                    If Cabecera.Lines(nNumLin).positive = 1 Then
            '                         Cabecera.Lines(nNumLin).FactorPeso = 1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 0
            '                    Else
            '                         Cabecera.Lines(nNumLin).FactorPeso = -1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 2
            '                    End If
                                'STP (pesado o no pesado)
                                If Cabecera.Lines(nNumLin).EsPesado Then
                                     'Cabecera.Lines(nNumLin).Weighted = 1
                                Else
                                     'Cabecera.Lines(nNumLin).Weighted = 0
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                                End If
                                
                                'CadenadeLog "paso 51"
                                ' STC (cancelado o no cancelado)
                                If Cabecera.Lines(nNumLin).LineaCancelada Then
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                                End If
                                'CadenadeLog "paso 52"
                                If MyRecAux.Fields("pricechanged") <> 0 Then
                                    Cabecera.Lines(nNumLin).SehaModificado = True
                                Else
                                    Cabecera.Lines(nNumLin).SehaModificado = False
                                End If
                                '¿? ModificadoTemporal As Boolean
                                'CadenadeLog "paso 53"
                                'CadenadeLog "8.8"
                                'c2f 1.1.1 aclarar ficha Vacuno...
                                'If IsNull(MyRecAux.Fields("eartag")) Or MyRecAux.Fields("eartag") = "" Then
                                    Cabecera.Lines(nNumLin).FichaVacuno = 0
                                'Else
                                '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("eartag")
                                'End If
                                'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                                'Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'CadenadeLog "8.9"
                                MyRecAux.Movenext
                            Loop
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            'CadenadeLog "paso 54"
                            
                            Actualiza_Tabla_Cabecera Mybase, Cabecera
                            
                            'CadenadeLog "paso 55"
                            Do_Events
                            If Cabecera.Incluir_en_Base = True Then
                                Actualiza_Tabla_Tickets Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtarti Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_GtVend Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtHora Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtsecs Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtInfVen Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Log Mybase, Cabecera
                            End If
                            
                            'CadenadeLog "paso 56"
                        End If
            
            
            
                        '''''''''''''''
                        'marcar...
                        'cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
                        'If nLR = 0 Then
                            cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
                        'Else
                        '    cm = "update pcscale.localhtickets set captured='*' where id=" & CStr(IdDoc)
                        'End If
                        CadenadeLogTotales "Marcando:" & "update mbd.hTickets set Captured='*' where id=" & CStr(IdDoc)

                        MyConnObj.Execute cm
                        ''''''''''
                        'CadenadeLog "paso 57"
                    Else 'REVISAR...
                    
                    
                            'CadenadeLog "paso 58"
                            nNumLin = 1
                            If MyRecSetTmp.Fields("saleform") = 1 Then
                                Cabecera.Lines(nNumLin).EsPesado = True
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            Else
                                Cabecera.Lines(nNumLin).EsPesado = False
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            End If
                            'CadenadeLog "paso 59"
                            'CadenadeLog "8.4"
                            
                            Cabecera.Lines(nNumLin).FactorPeso = 1
                            Cabecera.Lines(nNumLin).precio = MyRecSetTmp.Fields("price")
                            
                            'CadenadeLog "8.5"
                            'CadenadeLog "paso 60"
                            Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                            Cabecera.Lines(nNumLin).CodigoArticulo = MyRecSetTmp.Fields("item")
                            '3.4.4
                            'Cabecera.Lines(nNumLin).desc = MyRecAux.Fields("name")
                            ''''''
                            
                            'ver qué se hace IVA As Integer
                            Cabecera.Lines(nNumLin).IVA = 0
        
                            'CadenadeLog "paso 61"
                            'CadenadeLog "8.6"
                            
                            Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("linediscount") 'porcentaje
                            Cabecera.Lines(nNumLin).ImporteLinea = MyRecSetTmp.Fields("amount") 'MyRecSetTmp.Fields("importeconimpuestos") - MyRecSetTmp.Fields("dtoscabeceraconimpuestos")
                            Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecSetTmp.Fields("weight") * MyRecSetTmp.Fields("price"), 2) 'MyRecSetTmp.Fields("importe") - MyRecSetTmp.Fields("dtoscabecera")
                            'CadenadeLog "paso 62"
                            'CadenadeLog "8.7"
                            
                            'If MyRecSetTmp.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                            If MyRecSetTmp.Fields("linetype") = 0 Then
                                Cabecera.Lines(nNumLin).LineaCancelada = True
                            Else
                                Cabecera.Lines(nNumLin).LineaCancelada = False
                            End If
                            'CadenadeLog "paso 63"
                            'ver qué se hace con TipoLinea As Byte
                            'STV (positivo o negativo)
                            'Linea(Bucle).ModificadoTemporal = False
                            If MyRecSetTmp.Fields("amount") > 0 Then
                                'Cabecera.Lines(nNumLin).positive = 1
                                 Cabecera.Lines(nNumLin).FactorPeso = 1
                                 Cabecera.Lines(nNumLin).TipoLinea = 0
                            
                            Else
                                'Cabecera.Lines(nNumLin).positive = 0
                                 Cabecera.Lines(nNumLin).FactorPeso = -1
                                 Cabecera.Lines(nNumLin).TipoLinea = 2
                                
                            End If
                            'CadenadeLog "paso 64"
        '                    If Cabecera.Lines(nNumLin).positive = 1 Then
        '                         Cabecera.Lines(nNumLin).FactorPeso = 1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 0
        '                    Else
        '                         Cabecera.Lines(nNumLin).FactorPeso = -1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 2
        '                    End If
                            'STP (pesado o no pesado)
                            If Cabecera.Lines(nNumLin).EsPesado Then
                                 'Cabecera.Lines(nNumLin).Weighted = 1
                            Else
                                 'Cabecera.Lines(nNumLin).Weighted = 0
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                            End If
                            'CadenadeLog "paso 65"
                            
                            ' STC (cancelado o no cancelado)
                            If Cabecera.Lines(nNumLin).LineaCancelada Then
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                            End If
                            'CadenadeLog "paso 66"
                            'If MyRecSetTmp.Fields("pricechanged") <> 0 Then
                            '    Cabecera.Lines(nNumLin).SehaModificado = True
                            'Else
                                Cabecera.Lines(nNumLin).SehaModificado = False
                            'End If
                            '¿? ModificadoTemporal As Boolean
                            'CadenadeLog "paso 67"
                            'CadenadeLog "8.8"
                            'c2f 1.1.1 aclarar ficha Vacuno...
                            'If IsNull(MyRecSetTmp.Fields("eartag")) Or MyRecSetTmp.Fields("eartag") = "" Then
                                Cabecera.Lines(nNumLin).FichaVacuno = 0
                            'Else
                            '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecSetTmp.Fields("eartag")
                            'End If
                            'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                            'Cabecera.descuento = Cabecera.descuento + MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'CadenadeLog "8.9"
                    
                        'CadenadeLog "paso 68"
                        Actualiza_Tabla_Cabecera Mybase, Cabecera
                        Do_Events
                        If Cabecera.Incluir_en_Base = True Then
                            Actualiza_Tabla_Tickets Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Gtarti Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_GtVend Mybase, Cabecera
                            Do_Events
                            actualiza_tabla_gtHora Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Gtsecs Mybase, Cabecera
                            Do_Events
                            actualiza_tabla_gtInfVen Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Log Mybase, Cabecera
                        End If
                        'CadenadeLog "paso 69"
                        'marcar...
                        If nLR = 0 Then
                            If Tabla = 1 Then
                                cm = "update pcscale.packetlabels set customer='9999999999' where id=" & CStr(IdDoc)
                            Else
                                cm = "update pcscale.sservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                            End If
                        Else
                            If Tabla = 1 Then
                                cm = "update pcscale.localpacketlabels set customer='9999999999' where id=" & CStr(IdDoc)
                            Else
                                cm = "update pcscale.localsservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                            End If
                        End If
                        MyConnObj.Execute cm
                        ''''''''''
                        'CadenadeLog "paso 70"
                    End If
                    
                    MyRecSetTmp.Update
        
                    MyRecSetTmp.Movenext
                
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirLeeTique
                    Do_Events
                    
                Loop
            End If
            'CadenadeLog "paso 71"
            'On Error Resume Next
            If MyRecSetTmp.State <> 0 Then
                MyRecSetTmp.Close
                Set MyRecSetTmp = Nothing
            End If
            'On Error GoTo 0
SigueNlr:
            'CadenadeLog "paso 72"
            Next nLR
            
            End If
            'CadenadeLog "paso 73"
            MisEqp.Movenext
            
        Loop
        
        MisEqp.Close
        Set MisEqp = Nothing
        
    End If
    
SalirLeeTique:
''''''''
    On Error GoTo 0
    On Error GoTo fin
    'myRec.Close
    'Set myRec = Nothing
    If MyRecSetTmp.State <> 0 Then
        MyRecSetTmp.Close
        Set MyRecSetTmp = Nothing
    End If
    If Mybase.name <> "" Then
    Mybase.Close
    Set Mybase = Nothing
    End If
fin:
    If Err.Number <> 0 Then
        CadenadeLog "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
        On Error GoTo 0
        On Error GoTo FIN1
        If MyConnObj.State <> 0 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            Do_Events
            Sleep (200)
            frmControl.Switch_mode 0
            Do_Events
            Sleep (200)
            EstadoBal = 1000
        End If
    End If
FIN1:
    On Error GoTo 0

End Sub

'DROP TABLE IF EXISTS `pcscale`.`counters`;
'CREATE TABLE  `pcscale`.`counters` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Description` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaDesSec()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nResp As Integer

'    nResp = ElPing(sAddIpTouch)
'    If nResp = 1 Then

    If lTransPlu Then Exit Sub

    On Error GoTo FINDESSEC
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select * from seccion where borrado=false")
    On Error GoTo FINDESSEC
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            MyRecDes.Open "Select * from pcscale.counters where code=" & CStr(MyRecVar.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            frmMonitor.MostrarInfo CargaCadena(1417) & " " & CStr(MyRecVar.Fields("secc_maqui"))
            MyRecDes.Fields("code") = MyRecVar.Fields("secc_maqui")
            If Not IsNull(MyRecVar.Fields("nombre")) Then
                MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nombre")))
            Else
                MyRecDes.Fields("name") = ConvertStringToUtf8Bytes("Sección: " & CStr(MyRecVar.Fields("secc_maqui")))
            End If
            If Not IsNull(MyRecVar.Fields("descripcio")) Then
                MyRecDes.Fields("description") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("descripcio")))
            Else
                MyRecDes.Fields("description") = ConvertStringToUtf8Bytes("---" & CargaCadena(47) & " " & CStr(MyRecVar.Fields("secc_maqui")) & "---")
            End If
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirDesSec
            
        Loop
    
    
    End If
SalirDesSec:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINDESSEC:
    CadenadeLog "Error en MUE_EnviaDesSec:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
'    Else
'        canCelar = 1
'    End If
End Sub

'La idea de almacenamiento de datos se ha cambiado... Aquí fugura la Original.
' Como queda ahora sería la misma para Adaptador de protocolo, Alcampo y Sabeco...
'DROP TABLE IF EXISTS `pcscale`.`trademark`;
'CREATE TABLE  `pcscale`.`trademark` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `CO_Name` varchar(80) NOT NULL,          --> cab 1
'  `Description` varchar(80) NOT NULL,      --> cab 2
'  `Business` varchar(80) NOT NULL,         --> cab 3
'  `Slogan` varchar(80) NOT NULL,           --> cab 4
'  `Headquarter` varchar(80) NOT NULL,      --> cab 5
'  `CO_WEB` varchar(80) NOT NULL,           --> ley 1
'  `Shop_Name` varchar(80) NOT NULL,        --> ley 2
'  `Address` varchar(80) NOT NULL,          --> ley 3
'  `City` varchar(80) NOT NULL,             --> ley 4
'  `Phone` varchar(80) NOT NULL,            --> ley 5
'-------Solo Cabecera y Leyenda
'  `TIN` varchar(80) NOT NULL,              --> TxtF 1
'  `Shop_WEB` varchar(80) NOT NULL,         --> txtF 2
'  `Text1` varchar(80) NOT NULL,            --> txtF 3
'  `Text2` varchar(80) NOT NULL,            --> txtF 4
'  `Text3` varchar(80) NOT NULL,            --> txtF 5
'  `Text4` varchar(80) NOT NULL,            --> txtF 6
'  `Text5` varchar(80) NOT NULL,            --> txtF 7
'  `Text6` varchar(80) NOT NULL,            --> txtF 8
'''''''''''''''''''''''''''''''''
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaCabLey()
Dim Mybase As dao.Database
Dim MyRecVar As dao.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nResp As Integer
'    nResp = ElPing(sAddIpTouch)
'    If nResp = 1 Then

    On Error Resume Next
    Set Mybase = dao.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCABLEY
    End If
    On Error GoTo 0
    
    Set MyRecVar = Mybase.OpenRecordset("select * from trade")
    If Not MyRecVar.EOF Then
        'frmmonitor.barra1.Value = 0
        'frmmonitor.barra1.Max = MyRecVar.Recordcount
        'MyRecVar.MoveFirst
        'Do While Not MyRecVar.EOF
            
            MyRecDes.Open "Select * from pcscale.trademark", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                'Call frmControl.Balanza_ComError(0)
                GoTo FINCABLEY
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
    
'    MyReg.Fields("company") = Trim(TxtCab(0).TexT)
'    MyReg.Fields("name") = Trim(TxtCab(1).TexT)
'    MyReg.Fields("address") = Trim(TxtCab(2).TexT)
'
'    MyReg.Fields("quarter") = Trim(TxtLey(0).TexT)
'    MyReg.Fields("web") = Trim(TxtLey(1).TexT)
'    MyReg.Fields("shopweb") = Trim(TxtLey(2).TexT)
'    MyReg.Fields("description") = Trim(TxtLey(3).TexT)
'    MyReg.Fields("slogan") = Trim(TxtLey(4).TexT)
            
            frmMonitor.MostrarInfo CargaCadena(1418)
            MyRecDes.Fields("shop_name") = ConvertStringToUtf8Bytes(MyRecVar.Fields("company"))
            MyRecDes.Fields("address") = ConvertStringToUtf8Bytes(MyRecVar.Fields("name"))
            MyRecDes.Fields("city") = ConvertStringToUtf8Bytes(MyRecVar.Fields("address"))
            
            MyRecDes.Fields("text1") = ConvertStringToUtf8Bytes(MyRecVar.Fields("quarter"))
            MyRecDes.Fields("text2") = ConvertStringToUtf8Bytes(MyRecVar.Fields("web"))
            MyRecDes.Fields("text3") = ConvertStringToUtf8Bytes(MyRecVar.Fields("shopweb"))
            MyRecDes.Fields("text4") = ConvertStringToUtf8Bytes(MyRecVar.Fields("description"))
            MyRecDes.Fields("text5") = ConvertStringToUtf8Bytes(MyRecVar.Fields("slogan"))
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            'MyRecVar.Movenext
            
            'Call MonitorBarra
            'Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCabLey
            
        'Loop
    
    
    End If
SalirCabLey:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCABLEY:
    CadenadeLog "Error en MUE_EnviaCabLey:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
'    Else
'        canCelar = 1
'    End If
End Sub

'***********************************
'////////////////////////////
'TouchScale --> BASE MAESTRA
Public Sub AN_BaseMaestra(Modificaciones As Boolean, Optional ByVal lC As Boolean, Optional ByVal nEnv As Integer)
    Dim bucle As Integer
    Dim Base As dao.Database
    Dim SeBorra As Boolean
    Dim Beco As Long
    Dim Registro As dao.Recordset
    Dim nEnviados As Integer
    
    On Error GoTo fin
    If Trim(Command) = "" Then
        frmMonitor.Command1.Enabled = False
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = False
        frmMonitor.cmdmonitorstop.Enabled = True
        frmMonitor.Label1 = ""
        MostrarMonitor
    End If
    
    'nEnviados = 0
    
    lCogeTiquet = False
    canCelar = 0
    CancelHumano = False
    
    
    
    MUE_LanzaBaseMaestra Modificaciones, nEnv
    nEnviados = 2

    '*****************************
    ' borrado de datos eliminados
    'SE REALIZARÁN EN ENVIO...
    '*****************************
    'Set Base = OpenDatabase(Base_General) '1.7.2 OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo where borrado=true and tran_plu='*' and tran_pluSC10='*' and tran_en='*' and tran_el='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If (Not .EOF) Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from publi where borrado=true and tran_pub='*' and tran_pubSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from cabley where borrado=true and tran_cab='*' and tran_cabSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from codbar where borrado=true and tran_bar='*' and tran_barSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from vendedor where borrado=true and tran_vend='*' and tran_vendSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from teclas where borrado=true")
    ''2.0.30 ("select * from teclas where borrado=true and tran_tec='*' and tran_tecSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
fin:
On Error GoTo 0

    On Error Resume Next
    'Registro.Close
    '1.7.2 CerrarBase Base
    Set Base = Nothing
    'Workspaces(0).close
    
    If Trim(Command) = "" Then
        frmMonitor.Label1.Caption = CargaCadena(445)  '"Comunicación Finalizada"
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
    End If
    On Error GoTo 0
    lCogeTiquet = True
End Sub

'TouchScale -->MUE_LanzaBaseMaestra
'REVISADA VER COMO SE REALIZA FINALMENTE VACUNO...
Public Function MUE_LanzaBaseMaestra(Modificaciones As Boolean, Optional ByVal nEnv As Integer) As Boolean
'************************************************************************************
' envío de base maestra, realizando un ciclo por todas los sistemas existentes
'************************************************************************************
    Dim bucle As Integer
    Dim Marcar As Boolean
    Dim ErrorPrevio As Boolean
    Dim cm As String
    
    ErrorPrevio = False
    Marcar = False
    
    MUE_Sistemas
    
    For bucle = 0 To SisEur.NSistemas - 1
        If MultiEuroscale Then
            frmControl.Switch_mode 700 + bucle
        Else
            frmControl.Switch_mode 0
        End If
        'If Bucle = SisEur.NSistemas - 1 Then
        '    If Not ErrorPrevio Then Marcar = True
        'End If
        If CancelHumano Then Exit For
        canCelar = 0
        'touchScale Barna...
        'Incluir envio de Secciones (No Existe)
        'Incluir envio de Balanzas (No Existe)
        'Incluir envio de IVAS (No Existe)
        'cerrar conexión
               
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaDesSec
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        '1.1.3
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 7) Then MUE_EnviaArticuloNW bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        '''''''
        'Buscar "adaptar touchscale barna"
        'MUE_EnviaDesSec
        'MUE_EnviaCabLey
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaCabLey
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        ''''''''''''''''''''''''''''''''''
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaVendedor bucle, Modificaciones
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 2) Then MUE_EnviaCLientes bucle, Modificaciones, Marcar
        'se envian en artículos... If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 3) Then MUE_EnviaIVAS Bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 4) Then MUE_EnviaCodigosBarras bucle, Modificaciones
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 5) Then MUE_EnviarFichasVacuno bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 Then Call capturaPaises
        '1.1.3 If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 7) Then MUE_EnviaArticuloNW Bucle, Modificaciones, Marcar
        'cerrar conexión
        On Error Resume Next
        If MyConnObj.Errors.Count <> 0 Or MyConnObj.State <> 1 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            EstadoBal = 1000
        End If
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
        frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
        If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
        On Error Resume Next
        
        MyConnObj.Close
        Set MyConnObj = Nothing
        On Error GoTo 0
        EstadoBal = 1000
    
    
    Next bucle
    On Error GoTo 0
    
End Function

Public Sub KillTECLAS()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset

    frmControl.Switch_mode 0
    If EstadoBal = 0 Then


        On Error Resume Next
        CadenadeLog "pcscale.plukeys"
        cm = "delete from pcscale.plukeys"
        MyConnObj.Execute cm
        If Err.Number = 0 Then
            MsgBox "TECLAS BORRADAS.", vbExclamation
        End If
        On Error GoTo 0
        
        CadenadeLog "BORRAR TECLAS..."


    End If
    
End Sub

Public Sub KillARTICLES()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
Dim a As Variant

    a = MsgBox("EL BORRADO DE ARTÍCULOS CONLLEVA LA CAPTURA Y BORRADO DE TOTALES, ¿DESEA CONTINUAR?", vbYesNo)
    If a = vbYes Then
        frmControl.Switch_mode 0
        If EstadoBal = 0 Then
    
            CancelHumano = False
            frmMonitor.LblError = ""
            frmMonitor.Show
            frmMonitor.SetFocus
            
           '''''''
            'Call capturaClientes
            ''''''''
            LeeTickets 0
            LeeTickets 1, True
            LeeTickets 2, True
            
            If canCelar = 0 Then MUE_SolicitaPedido 0, False
            If canCelar = 0 Then MUE_SolicitaInventario 0, False
            If GA_Reinit Then
               Capa1A_GA_Erase_Totals (0)
            Else
               Capa1A_GA_Erase_Totals (1)
            End If
            If canCelar = 0 Then
                On Error Resume Next
                CadenadeLog "pcscale.plukeys"
                cm = "delete from pcscale.plukeys"
                MyConnObj.Execute cm
                On Error GoTo 0
        
        
                On Error Resume Next
                CadenadeLog "pcscale.items"
                cm = "delete from pcscale.items where code<>0"
                MyConnObj.Execute cm
                If Err.Number = 0 Then
                    MsgBox "ARTICULOS Y TECLAS BORRADOS.", vbExclamation
                End If
                On Error GoTo 0
                
                CadenadeLog "BORRAR ARTICULOS..."
                
                frmMonitor.Hide
                frmEpelsa.SetFocus
                
            End If
            
        End If
    End If
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////
'Adaptar TouchScale Barna
Public Sub BorraSrv()
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
On Error Resume Next
    
    CadenadeLog "BORRASRV..."
    
    On Error Resume Next
    CadenadeLog "pcscale.Barcodeean13"
    cm = "delete from pcscale.Barcodeean13"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.Bovine"
    cm = "delete from pcscale.Bovine"
    MyConnObj.Execute cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "pcscale.customers"
    cm = "delete from pcscale.customers"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.families"
    cm = "delete from pcscale.families"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.htickets"
    cm = "delete from pcscale.htickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.items"
    cm = "delete from pcscale.items where code<>0"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localerasetickets"
    cm = "delete from pcscale.localerasetickets"
    MyConnObj.Execute cm
    On Error GoTo 0
      
    On Error Resume Next
    CadenadeLog "pcscale.localhtickets"
    cm = "delete from pcscale.localhtickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localtickets"
    cm = "delete from pcscale.localtickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpacketlabels"
    cm = "delete from pcscale.localpacketlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpacketpresets"
    cm = "delete from pcscale.localpacketpresets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpaytickets"
    cm = "delete from pcscale.localpaytickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localplukeyschanges"
    cm = "delete from pcscale.localplukeyschanges"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localsservicelabels"
    cm = "delete from pcscale.localsservicelabels"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localvendorchanges"
    cm = "delete from pcscale.localvendorchanges"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localvendorpacket"
    cm = "delete from pcscale.localvendorpacket"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.ltickets"
    cm = "delete from pcscale.ltickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    
    'c2f 09-08-2012
    'If ReiniciarNum <> 0 Then
        On Error Resume Next
        CadenadeLog "pcscale.erasetickets"
        cm = "delete from pcscale.erasetickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        
        On Error Resume Next
        CadenadeLog "pcscale.localnumtickets"
        cm = "delete from pcscale.localnumtickets"
        MyConnObj.Execute cm
        On Error GoTo 0

        On Error Resume Next
        CadenadeLog "pcscale.numtickets"
        cm = "delete from pcscale.numtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
    'End If
    ''''''''''''''''''
    
    'c2f 09-08-2012 esta tabla es de configuración no alterar...
    'On Error Resume Next
    'CadenadeLog "pcscale.numberingtickets"
    'cm = "delete from pcscale.numberingtickets"
    'MyConnObj.Execute cm
    'On Error GoTo 0
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    CadenadeLog "pcscale.packetlabels"
    cm = "delete from pcscale.packetlabels"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.paymethods"
    cm = "delete from pcscale.paymethods"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.paytickets"
    cm = "delete from pcscale.paytickets"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.plukeys"
    cm = "delete from pcscale.plukeys"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.qlevels"
    cm = "delete from pcscale.qlevels"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.sservicelabels"
    cm = "delete from pcscale.sservicelabels"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.trademark"
    cm = "delete from pcscale.trademark"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.vendorkeys"
    cm = "delete from pcscale.vendorkeys"
    MyConnObj.Execute cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.vendors"
    cm = "delete from pcscale.vendors"
    MyConnObj.Execute cm
    On Error GoTo 0

   ' On Error Resume Next
   ' cm = "replace INTO pcscale.Trademark VALUES(1,'','','','','','','','','','','','','','','','','','');"
   ' MyConnObj.Execute cm
   ' If Err.Number <> 0 Then
   '     On Error GoTo 0
   '     cm = "replace INTO pcscale.Trademark VALUES(1,'','','','','','','','','','','','','','','','','','','','','','','','','','','','');"
   '     MyConnObj.Execute cm
   '
   ''
    'End If
    'On Error GoTo 0
    'c2f 09-08-2012 --> Ver con corrección Balanza
    ' En principio dejo los código de barras necesarios Prueba ECI para el equipo 15
    ' ¡¡¡ QUITAR MÁS ADELANTE Y VER QUÉ CÓDIGOS DE BARRAS SE DEJAN POR DEFECTO...!!!
    'cm = "REPLACE INTO BarcodeEAN13 VALUES(1,'S',15, '2502215TTTTT', '26CCCCCiiiii', '2502215TTTTT', '26CCCCCiiiii', '26CCCCCiiiii', '25CCCCCiiiii', '2YTTTTTIIIII', '20VVViiiiiii', '10VVVIIIIIII', '2YTTTTTiiiii', '2CCCCCCiiiii', '2CCCCCCiiiii',  '2CCCCCCIIIII');"
    'MyConnObj.Execute cm
    '''''''
    
    
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
    End If

On Error GoTo 0

End Sub

Public Function ElPing(ElHost As String) As Integer
    '///////////////////////////////////////////////
    '//  Lanza ping a la dirección que se indique //
    '//  llamando al OCX de ping                  //
    '///////////////////////////////////////////////
    Dim Contador As Integer
    Dim Echo As ICMP_ECHO_REPLY
    Dim Resp As Long
    If (Dir(App.Path & "\noping.txt") <> "" Or Dir(App.Path & "\noping") <> "") Then
        ElPing = 1
    Else
        Resp = ICMP_Ping(ElHost, "alive", Echo)
        If Resp = 0 Then
            ElPing = 1
        Else
            ElPing = 0
        End If
    End If
End Function

'''''''''''''''''''''
'''STOCK + INVENTARIO
'''''''''''''''''''''
'DROP TABLE IF EXISTS `pcscale`.`stock`;
'CREATE TABLE  `pcscale`.`stock` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Inventory_Weight` decimal(12,3) DEFAULT '0.000',
'  `Inventory_Units` int(11) DEFAULT '0',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `stock_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
''''''''''''''''''''''
'''PEDIDO
''''''''''''''''''''''
'DROP TABLE IF EXISTS `pcscale`.`shoporder`;
'CREATE TABLE  `pcscale`.`shoporder` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Observations` varchar(127) DEFAULT NULL,
'  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `shoporder_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;

Public Sub MUE_SolicitaPedido(NSistema As Integer, FinalDia As Boolean)
    Dim RutaPedido As String
    Dim RutaStock As String
    'Dim Mybase As DAO.Database
    'Dim MyRecVar As DAO.Recordset
    Dim nf As Integer
    'Dim nG As Integer
    Dim sF As String
    'Dim sG As String
    On Error GoTo PedFin
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        If Not HayMulti Then
            RutaPedido = App.Path & "\01ped.sdf"
            'RutaPedido = App.Path & "\01sto.sdf"
        Else
            If TiendaActual < 100 Then
                RutaPedido = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "ped.sdf" '& Format(nsistema, "00")
                'RutaStock = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "sto.sdf." & Format(nsistema, "00")
            Else
                RutaPedido = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "ped.sdf" '& Format(nsistema, "000")
                'RutaStock = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "sto.sdf." & Format(nsistema, "000")
            End If
        End If
        frmMonitor.MostrarInfo CargaCadena(1106)
        'If Dir(App.Path & "\pedplusec.txt") <> "" Then
        '    Resp = frmControl.Balanza.Query_Order(RutaPedido, TiendaActual)
        'Else
        '    Resp = frmControl.Balanza.Query_Order(RutaPedido)
        'End If
        'Set Mybase = DAO.OpenDatabase(Base_General)
        'MyRecSet.CursorLocation = adUseClient
        
        MyRecSet.Open "Select * from pcscale.shoporder where code<>0", MyConnObj, adOpenKeyset, adLockOptimistic '1.8.9 instr(lcase(pcscale.shoporder.observations),'(captured)')=0", MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecSet.EOF Then
            nf = FreeFile()
            '1.8.9 Open RutaPedido For Append As #nf
            Open RutaPedido For Output As #nf
            'nG = FreeFile()
            'Open RutaStock For Output As #nG
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                sF = Format(MyRecSet.Fields("code"), "000000")
                If MyRecSet.Fields("saleform") <> 0 Then
                    sF = sF & "W"
                    sF = sF & Format(MyRecSet.Fields("weight") * 1000, "00000000")
                Else
                    sF = sF & "U"
                    sF = sF & Format(MyRecSet.Fields("units"), "00000000")
                End If
                'sG = sF
                sF = sF & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                If Len(sF) < 40 Then
                    sF = sF & Mid("                                        ", 1, 40 - Len(sF))
                End If
                sF = sF & Format(MyRecSet.Fields("dattim"), "dd-mm-yy hh:mm:ss")
                Print #nf, sF
                '1.8.9 MyRecSet.Fields("observations") = Mid(MyRecSet.Fields("observations"), 1, 110) & "(captured)"
                '1.8.9 MyRecSet.Update
                MyRecSet.Movenext
            Loop
            Close #nf
        
        End If
        MyRecSet.Close
        Set MyRecSet = Nothing
        
        
        
        'If canCelar = 0 Then
        '    If (FinalDia And BorrarPedidoFin) Or BorrarPedido Then
        '        frmMonitor.MostrarInfo CargaCadena(1108)
        '        frmControl.Balanza.Erase_Order
        '    End If
        '
        'End If
        frmMonitor.Label1.Caption = CargaCadena(445)
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = True
    End If
    If canCelar = 0 Then
        frmMonitor.MostrarDato CargaCadena(1096) & " " & RutaPedido
    Else
        On Error Resume Next
        'Kill RutaPedido
    End If
PedFin:
End Sub

Public Sub MUE_SolicitaInventario(NSistema As Integer, FinalDia As Boolean)
    Dim RutaInventario As String
    Dim RutaStock As String
    Dim nf As Integer
    Dim nG As Integer
    Dim sF As String
    Dim sG As String
    Dim cm As String
    
    On Error GoTo PedFin
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        If Not HayMulti Then
            RutaInventario = App.Path & "\01inv.sdf"
            RutaStock = App.Path & "\01sto.sdf"
        Else
            If TiendaActual < 100 Then
                RutaInventario = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "inv.sdf" '& Format(nsistema, "00")
                RutaStock = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "sto.sdf" '& Format(nsistema, "00")
            Else
                RutaInventario = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "inv.sdf" '& Format(nsistema, "000")
                RutaStock = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "sto.sdf" '& Format(nsistema, "000")
            End If
        End If
        frmMonitor.MostrarInfo "Inventory & Stock."
'DROP TABLE IF EXISTS `pcscale`.`stock`;
'CREATE TABLE  `pcscale`.`stock` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `Weight` decimal(12,3) DEFAULT '0.000',
'  `Units` int(11) DEFAULT '0',
'  `Inventory_Weight` decimal(12,3) DEFAULT '0.000',
'  `Inventory_Units` int(11) DEFAULT '0',
'  PRIMARY KEY (`Code`),
'  CONSTRAINT `stock_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
        MyRecSet.Open "Select * from pcscale.stock where code<>0", MyConnObj, adOpenKeyset, adLockOptimistic '1.8.9 instr(lcase(pcscale.stock.name),'(captured)')=0", MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecSet.EOF Then
            nf = FreeFile()
            '1.8.9 Open RutaInventario For Append As #nf
            Open RutaInventario For Output As #nf
            nG = FreeFile()
            '1.8.9 Open RutaStock For Append As #nG
            Open RutaStock For Output As #nG
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                sF = Format(MyRecSet.Fields("code"), "000000")
                sG = sF
                If MyRecSet.Fields("saleform") <> 0 Then
                    sG = sG & "W"
                    sG = sG & Format(MyRecSet.Fields("weight") * 1000, "00000000")
                    sF = sF & "W"
                    sF = sF & Format(MyRecSet.Fields("inventory_weight") * 1000, "00000000")
                Else
                    sG = sG & "U"
                    sG = sG & Format(MyRecSet.Fields("units"), "00000000")
                    sF = sF & "U"
                    sF = sF & Format(MyRecSet.Fields("inventory_units"), "00000000") 'inventory_weight, inventory_units
                End If
                
                sF = sF & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                sG = sG & Mid(Trim(MyRecSet.Fields("name")), 1, 25)
                
                If Len(sF) < 40 Then
                    sF = sF & Mid("                                        ", 1, 40 - Len(sF))
                End If
                If Len(sG) < 40 Then
                    sG = sG & Mid("                                        ", 1, 40 - Len(sG))
                End If
                
                sF = sF & Format(Now, "dd-mm-yy hh:mm:ss")
                sG = sG & Format(Now, "dd-mm-yy hh:mm:ss")
                
                Print #nf, sF
                Print #nG, sG
                
                'MyRecSet.Fields("name") = Mid(MyRecSet.Fields("name"), 1, 110) & "(captured)"
                
                'MyRecSet.Update
                MyRecSet.Movenext
                
            Loop
            Close #nf
            Close #nG
        End If
        MyRecSet.Close
        Set MyRecSet = Nothing
        
        frmMonitor.Label1.Caption = CargaCadena(445)
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = True
    
    End If
    If canCelar = 0 Then
        frmMonitor.MostrarDato "Inventory:" & " " & RutaInventario
        frmMonitor.MostrarDato "Stock:" & " " & RutaStock
        If BorrarInventario Then '1.9.1 j.mer.
            On Error GoTo 0
            On Error Resume Next
            'inventory_weight, inventory_units
            cm = "update pcscale.stock set inventory_weight=0 where code<>0"
            MyConnObj.Execute cm
            ''''''''''''''''''''''''''''''''
            On Error GoTo 0
            On Error Resume Next
            'inventory_weight, inventory_units
            cm = "update pcscale.stock set inventory_units=0 where code<>0"
            MyConnObj.Execute cm
            ''''''''''''''''''''''''''''''''
            On Error GoTo 0
            On Error Resume Next
        End If
    Else
        On Error Resume Next
    End If
PedFin:
End Sub

'Public Sub AN_PedidoInventario()
'    Call MUE_SolicitaPedido
'    Call MUE_SolicitaInventario
'End Sub

'Public Function Capa1B_Price_T(myprice As Epel_Price_Trm)
'    Dim MyRecAux As New ADODB.Recordset
'    Dim cm As String
'    Dim nF As Integer
'    Dim nC As Long
'    Dim ncont As Integer
'    Dim Resp As Long
'    Dim nPrice As Double
''Public Type Epel_Price_Trm
''    Code As Long
''    Price1 As Long
''    Weight1 As Long
''    Price2 As Long
''    Weight2 As Long
''    lErase As Integer
''End Type
''DROP TABLE IF EXISTS `pcscale`.`itempricerates`;
''CREATE TABLE  `pcscale`.`itempricerates` (
''  `Id` int(11) NOT NULL AUTO_INCREMENT,
''  `Code` int(11) NOT NULL,
''  `Rate` decimal(1,0) NOT NULL DEFAULT '1',
''  `Price` decimal(8,2) NOT NULL DEFAULT '0.00',
''  `Step` decimal(5,3) NOT NULL DEFAULT '0.000',
''  PRIMARY KEY (`Id`),
''  UNIQUE KEY `idx_rate` (`Code`,`Rate`,`Step`),
''  KEY `Rate` (`Rate`),
''  CONSTRAINT `itempricerates_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE,
''  CONSTRAINT `itempricerates_ibfk_2` FOREIGN KEY (`Rate`) REFERENCES `pricerates` (`Code`) ON DELETE CASCADE
'') ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
'CREATE TABLE `pricerates` (
'  `Code` decimal(1,0) NOT NULL,
'  `Name` varchar(64) NOT NULL DEFAULT '',
'  PRIMARY KEY (`Code`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
'    Dim lC
'    On Error GoTo FIN_Capa1B_Price_T
'    cm = "delete from pcscale.itempricerates where code=" & CStr(myprice.Code)
'    MyConnObj.Execute cm
'
'    If myprice.lErase = 0 Then
'
'        On Error Resume Next
'        MyRecAux.Open "select * from pcscale.itempricerates where code=" & CStr(myprice.Code), MyConnObj, adOpenKeyset, adLockOptimistic
'        If Err.Number <> 0 Then
'            Resp = 1000
'            GoTo FIN_Capa1B_Price_T
'        End If
'        On Error GoTo FIN_Capa1B_Price_T
'        If (myprice.Price1 <> 0) And (myprice.Weight1 <> 0) Then
'            MyRecAux.AddNew
'            MyRecAux.Fields("code") = myprice.Code
'            MyRecAux.Fields("rate") = 0
'            MyRecAux.Fields("price") = myprice.Price1 / 100
'            MyRecAux.Fields("step") = myprice.Weight1 / 1000
'        End If
'        If (myprice.Price2 <> 0) And (myprice.Weight2 <> 0) Then
'            MyRecAux.AddNew
'            MyRecAux.Fields("code") = myprice.Code
'            MyRecAux.Fields("rate") = 0
'            MyRecAux.Fields("price") = myprice.Price2 / 100
'            MyRecAux.Fields("step") = myprice.Weight2 / 1000
'        End If
'
'        MyRecAux.Update
'        MyRecAux.Close
'        Set MyRecAux = Nothing'
'
'    End If
'
'    Resp = 0'
'
'FIN_Capa1B_Price_T:
'    If Err.Number <> 0 Then
'        CadenadeLog "Error en Price_T. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
'        On Error GoTo 0
'        Resp = 1004
'    End If'
'
'    PonDato " Cod: " & myprice.Code
'    Capa1B_Price_T = Resp'
'
'End Function

#End If

Public Function getNtotalTS(ByVal sfecha As Variant) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    On Error Resume Next
    'frmControl.Switch_mode 0
    MyRecAux.CursorLocation = adUseClient
    MyRecAux.Open "Select * from pcscale.totaldaily where date='" & Format(CDate(sfecha), "yyyy-mm-dd") & "'", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
        Resp = 0
    Else
        If MyRecAux.EOF Then
            Resp = 0
        Else
            Resp = MyRecAux.Fields("number")
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    'MyConnObj.Close
    'Set MyConnObj = Nothing
    getNtotalTS = Resp
End Function

'''Prueba copia BBDD mysql
Public Sub cmcSchemaADO_Click()
    Dim strSource As String
    'Dim cnnSrc As New ADODB.Connection
    Dim rstSrc As New ADODB.Recordset
    Dim catSrc As New adox.Catalog
    Dim tblSrc As New adox.Table
    Dim colSrc As New adox.Column
    Dim ixSrc As New adox.Index
    Dim propSrc As Property
    Dim iCount As Integer
    Dim iCt As Integer
    Dim iCt2 As Integer
    Dim nodX As Node
    Dim nf As Integer
    Dim Tabla As TableDef
    Dim Base As dao.Database
    Dim Registro As ADODB.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Do_Events
    
    'Set Tabla = Base.CreateTableDef("gen_sam")
    
'    With cdlgFile
'        .DialogTitle = "Select Database to Copy"
'        .Filter = "Access 97/2000 (*.mdb)|*.mdb"
'        .CancelError = True
'        .InitDir = "C:\My Documents"
'        .ShowOpen
'    End With
'    strSource = cdlgFile.FileName
    
'    cnnSrc.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSource
'    Set catSrc.ActiveConnection = cnnSrc
    
frmControl.Switch_mode 0
Set catSrc.ActiveConnection = MyConnObj
Set Registro = MyConnObj.OpenSchema(adSchemaTables)



    'Set nodX = TreeView1.Nodes.Add(, , "R", strSource)
    'iCount = 1
    'iCt = 1
    nf = FreeFile()
    Open App.Path & "\bbddtables.txt" For Output As #nf
    For Each tblSrc In catSrc.Tables
        If tblSrc.type = "TABLE" And InStr(1, tblSrc.name, "MSys", vbTextCompare) = 0 Then
            If LCase(tblSrc.name) = "htickets" Or LCase(tblSrc.name) = "ltickets" Then
                Print #nf, tblSrc.name
                Print #nf, "---------------------------"
                
                For Each colSrc In tblSrc.Columns
                    Print #nf, colSrc.name & ", " & colSrc.type & ", " & colSrc.DefinedSize
                Next colSrc
            
                
            End If
            
        End If
    Next tblSrc
    Close #nf
    
End Sub

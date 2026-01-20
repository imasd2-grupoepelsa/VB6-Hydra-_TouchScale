VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Ejemplo"
   ClientHeight    =   6225
   ClientLeft      =   2025
   ClientTop       =   1545
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9090
   Begin VB.CommandButton CmdTck 
      Caption         =   "Captura Tiques."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6240
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrmMain.frx":0000
      Top             =   1080
      Width           =   8895
   End
   Begin VB.CommandButton CmdErase 
      Caption         =   "Reiniciar Base de datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar Conexión"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Conectar con Base de datos Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Se supone instalado MySQL Connector 5.1.11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label Label1 
      Caption         =   "Se supone conexión contra host Local puerto 3306"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
    If MyConnObj.State <> 0 Then
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
End Sub

Private Sub CmdConnect_Click()
    If MyConnObj.State <> 0 Then
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
    MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & "127.0.0.1" & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
End Sub

Public Sub BorraSrv()
Dim Cm As String

    On Error GoTo ControlErr
    If MyConnObj.State = 1 Then

        Cm = "delete from pcscale.Barcodeean13"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.Bovine"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.customers"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.families"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.htickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.items"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localerasetickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localhtickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localtickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localpacketlabels"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localpacketpresets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localpaytickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localplukeyschanges"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localsservicelabels"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localvendorchanges"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.localvendorpacket"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.ltickets"
        MyConnObj.Execute Cm
        Cm = "truncate pcscale.erasetickets"
        MyConnObj.Execute Cm
        Cm = "truncate pcscale.localnumtickets"
        MyConnObj.Execute Cm
        Cm = "truncate pcscale.numtickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.packetlabels"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.paymethods"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.paytickets"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.plukeys"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.qlevels"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.sservicelabels"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.trademark"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.vendorkeys"
        MyConnObj.Execute Cm
        Cm = "delete from pcscale.vendors"
        MyConnObj.Execute Cm
    
        Cm = "REPLACE INTO Trademark VALUES(1,'','','','','','','','','','','','','','','','','','');"
        MyConnObj.Execute Cm
    
    Else
        MsgBox "NO EXISTE CONEXIÓN CON BASE DE DATOS...", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    MsgBox "Base de Datos Reiniciada...", vbExclamation
    Exit Sub
    
ControlErr:
    MsgBox "Error:" & CStr(Err.Number) & " Desc.:" & Err.Description
    On Error GoTo 0
End Sub

Private Sub CmdErase_Click()
    Call BorraSrv
End Sub

    
Public Sub LeeTickets()
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim Cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim ncont As Long
    Dim dLImporte As Double
    
    If MyConnObj.State = 1 Then
    
    On Error GoTo FIN
    
    MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and status='C' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSetTmp.EOF Then
        Do While Not MyRecSetTmp.EOF
            Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
            Fhora = Mid(Ffecha, 12, 8)
            Ffecha = Mid(Ffecha, 1, 10)
            IdDoc = MyRecSetTmp.Fields("id")
            ''''medio de pago
            '1=efectivo
            '2=t.crédito
            '3=cheque
            Cabecera.tipoPago = 1
            Cabecera.cantidadEntre = 0
            MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            '''''''''''''''''
            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
            dLImporte = MyRecAux.Fields("importe")
            MyRecAux.Close
            Set MyRecAux = Nothing
            'recoger datos
            Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
            Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
            Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
            Cabecera.NCliente = MyRecSetTmp.Fields("customer")
            Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.hora = Fhora
            Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
            Cabecera.ImporteTotal = dLImporte
            Cabecera.NLineas = nNumLin
            Cabecera.LineaInicial = 1
            Cabecera.LineaFinal = nNumLin
            Cabecera.LineasMensaje = nNumLin
            If MyRecSetTmp.Fields("tcktype") = "N" Then
                Cabecera.CadenaEstadoTiquet = "Local"
            Else
                Cabecera.CadenaEstadoTiquet = "Red"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            dLImporte = Round(MyRecAux.Fields("importe"), 2)
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.ImporteBruto = dLImporte
            Cabecera.Impuestos = 0
            Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal
            Cabecera.documentoabono = 0
            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.LineasCanceladas = nNumLin
            Select Case MyRecSetTmp.Fields("status")
            Case "C"
                If Cabecera.CadenaEstadoTiquet = "Local" Then
                        Cabecera.tipoTiquet = 4
                Else
                        Cabecera.tipoTiquet = 1
                End If
            End Select
            Cabecera.EstadoTiquet = 16
            Cabecera.Cambio = Cabecera.ImporteTotal - Cabecera.cantidadEntre
            Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")

            'LINEAS
            nNumLin = 0
            MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF
                    nNumLin = nNumLin + 1
                    If MyRecAux.Fields("saleform") = 1 Then
                        Cabecera.Lines(nNumLin).EsPesado = True
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                    Else
                        Cabecera.Lines(nNumLin).EsPesado = False
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                    End If
                    Cabecera.Lines(nNumLin).FactorPeso = 1
                    Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                    Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                    Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                    Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("linediscount") 'porcentaje
                    Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("amount")
                    Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2)
                    If MyRecAux.Fields("linetype") = 0 Then
                        Cabecera.Lines(nNumLin).LineaCancelada = True
                    Else
                        Cabecera.Lines(nNumLin).LineaCancelada = False
                    End If
                    Cabecera.Lines(nNumLin).SehaModificado = False
                    Cabecera.Lines(nNumLin).FichaVacuno = 0
                    MyRecAux.MoveNext
                Loop
                MyRecAux.Close
                Set MyRecAux = Nothing
            End If
            '''''''''''''''
            'marcar...
            Cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
            MyConnObj.Execute Cm
            ''''''''''
            MyRecSetTmp.Update
            
            MsgBox "Hay un Tique en la estructura...", vbExclamation
            
            MyRecSetTmp.MoveNext
            
        Loop
    End If

SalirLeeTique:
''''''''

    On Error Resume Next
    MyRecSetTmp.Close
    Set MyRecSetTmp = Nothing
    
FIN:
    If Err.Number <> 0 Then
        MsgBox "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
    End If
    On Error GoTo 0
    
    Else
        MsgBox "No existe conexión con Base de Datos...", vbExclamation
    End If
End Sub

Private Sub CmdTck_Click()
    Call LeeTickets
End Sub

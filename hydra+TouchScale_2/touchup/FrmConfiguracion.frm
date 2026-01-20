VERSION 5.00
Begin VB.Form FrmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexión"
   ClientHeight    =   4710
   ClientLeft      =   6960
   ClientTop       =   4395
   ClientWidth     =   4185
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdVerificaConexión 
      Caption         =   "Test Conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Acceso a todos los Artículos"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   1800
      TabIndex        =   5
      Text            =   "3306"
      Top             =   1440
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   408
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1572
   End
   Begin VB.Label Label4 
      Caption         =   "Sección Servidora:"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Ip Servidora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Puerto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca configuraciones solicitadas. (La Balanza Servidora será la Balanza número 1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DROP TABLE IF EXISTS `pcscale`.`scales`;
'CREATE TABLE  `pcscale`.`scales` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `MAC` varchar(17) NOT NULL,
'  `IP` varchar(16) NOT NULL,
'  `Number` int(11) NOT NULL DEFAULT '0',
'  `Counter` int(11) NOT NULL DEFAULT '0',
'  `Name` varchar(64) NOT NULL DEFAULT '',
'  `Type` char(1) DEFAULT NULL,
'  `AppVersion` varchar(16) DEFAULT '',
'  `SerialNumber` varchar(10) DEFAULT '',
'  `System` varchar(32) DEFAULT '',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `MAC` (`MAC`)
') ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=utf8;

'Public Sub CmdVerificaConexión_Click()
'    Dim myBase As DAO.Database
'    Dim myRS As DAO.Recordset
'    Dim cm As String
'    Dim vbYn As Variant
'    Dim MyRsDes As New ADODB.Recordset
'
'    lMySqlAbierta = False
'    If sIpAdd = "" Then
'        sIpAdd = toIpAdd(Trim(Text1.Text))
'        nIpPort = Val(Text2.Text)
'    End If
'    If Trim(Command) = "" Then
'        Call AbrirMySQL(toIpAdd(Trim(Text1.Text)))
'    Else
'        Call AbrirMySQL
'    End If
'    If lMySqlAbierta Then
'        If Trim(Command) = "" Then
'            FrmAviso.LblAviso = "Conexión Correcta con:" & sIpAdd & ":" & CStr(nIpPort)
'        End If
'    Else
'        If Trim(Command) = "" Then
'            FrmAviso.LblAviso = "Sin Conexión con:" & sIpAdd & ":" & CStr(nIpPort)
'        End If
'    End If
'    If Trim(Command) = "" Then
'        Load FrmAviso
'        FrmAviso.Show
'        Do While FrmAviso.Visible
'            Do_Events
'        Loop
'    End If
'    If lMySqlAbierta Then
'        If Trim(Command) = "" Then
'            vbYn = MsgBox("Por favor, confirme capturar Balanzas existentes en la red.", vbYesNo)
'        Else
'            vbYn = vbYes
'        End If
'        If vbYn = vbYes Then
'            Set myBase = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
'            cm = "delete from balanzas where scalenr<>1"
'            myBase.Execute cm
'            Set myRS = myBase.OpenRecordset("select * from balanzas where scalenr<>1")
'            cm = "select * from scales where type<>'S'"
'            MyRsDes.Open cm, DbMysqli, adOpenDynamic, adLockOptimistic
'            If Not MyRsDes.EOF Then
'                MyRsDes.MoveFirst
'                Do While Not MyRsDes.EOF
'                    myRS.AddNew
'                    myRS.Fields("scalenr") = MyRsDes.Fields("number")
'                    myRS.Fields("counter") = MyRsDes.Fields("counter")
'                    myRS.Fields("section") = MyRsDes.Fields("counter")
'                    myRS.Fields("ip") = GeneraTextoIP(MyRsDes.Fields("ip"))
'                    myRS.Update
'                    MyRsDes.MoveNext
'                Loop
'
'            End If
'            On Error Resume Next
'            myBase.Close
'            Set myBase = Nothing
'            myRS.Close
'            Set myRS = Nothing
'            MyRsDes.Close
'            Set MyRsDes = Nothing
'            On Error GoTo 0
'        End If
'    End If
'End Sub

'Configurar Sección y Dirección Ip Balanza
Private Sub Command1_Click(Index As Integer)
    Dim sData As String
    Dim Retval As Long
    Dim BufLabel As String
    Dim myBase As DAO.Database
    Dim myRS As DAO.Recordset
    Dim cm As String
    Dim lEsAlta As Boolean
    Dim lTmp As Boolean
    
    Me.Enabled = False
    Select Case Index
        Case 0
            sData = Text1.Text
            If GeneraTextoIP(sData) Then
                Text1.Text = sData
            Else
                MsgBox "IP incorrecta", vbCritical
                GoTo Salir
            End If
            On Error Resume Next
            Set myBase = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
            If Err.Number = 0 Then
                On Error GoTo 0
                Set myRS = myBase.OpenRecordset("select * from balanzas where port<>0")
                If Not myRS.EOF Then
                    myRS.Edit
                    If Trim(myRS.Fields("ip")) <> Trim(sData) Then
                        cm = "delete from balanzas"
                        myBase.Execute cm
                        lEsAlta = True
                        myRS.AddNew
                    End If
                Else
                    cm = "delete from balanzas"
                    myBase.Execute cm
                    lEsAlta = True
                    myRS.AddNew
                End If
                myRS.Fields("scalenr") = 1
                myRS.Fields("counter") = Val(Text3.Text)
                myRS.Fields("section") = Val(Text3.Text)
                myRS.Fields("ip") = sData
                myRS.Fields("port") = Val(Text2.Text)
                myRS.Fields("ItemImagePath") = App.Path & "\images"
                myRS.Fields("PublicityImagePath") = App.Path & "\images"
                myRS.Fields("QualityImagePath") = App.Path & "\images"
                myRS.Fields("allitems") = CBool(Check1.Value)
                If lEsAlta Then
                    myRS.Fields("crfknr") = 0
                    myRS.Fields("crfknm") = ""
                    myRS.Fields("Balknr") = 2
                    myRS.Fields("balknm") = "SS-2F-medium-num"
                End If
                myRS.Update
                ''''''''''''
                myRS.Close
                Set myRS = myBase.OpenRecordset("select * from balanzas where port<>0")
                myRS.Edit
                sIpAdd = Trim(myRS.Fields("ip"))
                nIpPort = CStr(myRS.Fields("port"))
                nSecSer = myRS.Fields("counter")
                
                lTmp = lAllItem
                
                If myRS.Fields("allitems") Then
                    lAllItem = True
                Else
                    lAllItem = False
                End If
                myRS.Close
                If lTmp <> lAllItem And lSalir = False Then
                    MsgBox "Se ha modificado tipo de acceso a Artículos. Se debe reiniciar utilidad."
                    End
                End If
                '''''''''''''
            Else
                MsgBox "No es posible abrir TOUCHSCALE.MDB", vbCritical
            End If
            On Error GoTo 0
            
            myBase.Close
            Unload Me
            Exit Sub
            
        Case 1
            Unload Me
            If lSalir Then End
            Exit Sub
            'End
    End Select
    
Salir:
    
    Me.Enabled = True
    
End Sub



Private Sub Form_Load()
    Dim myBase As DAO.Database
    Dim myRS As DAO.Recordset
    On Error Resume Next
    Set myBase = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    If Err.Number = 0 Then
        Set myRS = myBase.OpenRecordset("select * from balanzas where port<>0") 'solo debería ser la servidora
        If Not myRS.EOF Then
            myRS.Edit
            Me.Text1.Text = Trim(myRS.Fields("ip"))
            Me.Text2.Text = CStr(myRS.Fields("port"))
            Me.Text3.Text = CStr(myRS.Fields("counter"))
            If myRS.Fields("allitems") Then
                Check1.Value = vbChecked
            Else
                Check1.Value = vbUnchecked
            End If
        End If
        myRS.Close
    End If
    myBase.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case 48 To 59
        Case 46
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case 48 To 59
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case 48 To 59
        Case Else
            KeyAscii = 0
    End Select
End Sub



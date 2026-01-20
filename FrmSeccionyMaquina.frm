VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSeccionyMaquina 
   Caption         =   "Secciones y Equipos"
   ClientHeight    =   5820
   ClientLeft      =   6960
   ClientTop       =   2400
   ClientWidth     =   6330
   Icon            =   "FrmSeccionyMaquina.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6330
   Begin RichTextLib.RichTextBox info 
      Height          =   3255
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
      _Version        =   393217
      TextRTF         =   $"FrmSeccionyMaquina.frx":000C
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   1092
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
      Width           =   3852
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Command1"
         Height          =   372
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   3612
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Command1"
         Height          =   372
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   3612
      End
   End
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   2760
      Top             =   3840
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":0092
            Key             =   "pc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":01EE
            Key             =   "seccion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":034A
            Key             =   "v12"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":04A6
            Key             =   "redmarte"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":0600
            Key             =   "g386"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":075C
            Key             =   "g386M"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeccionyMaquina.frx":08B8
            Key             =   "iv4lcd"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Arbol 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   612
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   3852
   End
   Begin VB.Label lblinfo2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   3852
   End
   Begin VB.Menu mnuContextualBalanza 
      Caption         =   "Contextual"
      Begin VB.Menu mnuModificaBalanza 
         Caption         =   "Modificar Balanza"
      End
      Begin VB.Menu mnuEliminaBalanza 
         Caption         =   "Eliminar Balanza"
      End
   End
   Begin VB.Menu mnuContextualSeccion 
      Caption         =   "Contextual"
      Begin VB.Menu mnuModificaMostrador 
         Caption         =   "Modificar Mostrador"
      End
      Begin VB.Menu mnuEliminaMostrador 
         Caption         =   "Eliminar Mostrador"
      End
      Begin VB.Menu mnuVoid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreaBalanza 
         Caption         =   "Crear y eliminar Balanzas"
      End
   End
   Begin VB.Menu mnuContextualSistema 
      Caption         =   "Contextual"
      Begin VB.Menu mnuCreaSeccion 
         Caption         =   "Crear y Eliminar Mostradores"
      End
   End
End
Attribute VB_Name = "FrmSeccionyMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nodo As String
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long
Private Sub Me_Init()
    Dim Bucle As Long
    Me_Width = Me.Width
    Me_Height = Me.Height
    ReDim ct_top(Me.Controls.Count)
    ReDim CT_Left(Me.Controls.Count)
    ReDim CT_Width(Me.Controls.Count)
    ReDim CT_Height(Me.Controls.Count)
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        ct_top(Bucle) = Me.Controls(Bucle).top
        CT_Left(Bucle) = Me.Controls(Bucle).left
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        On Error GoTo 0
    Next Bucle
End Sub
Private Sub Me_Resize()
    Dim RelW As Double
    Dim Bucle As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        RelW = Me.Width / Me_Width
        Me.Controls(Bucle).Width = RelW * CT_Width(Bucle)
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        Me.Controls(Bucle).left = RelW * CT_Left(Bucle)
        CT_Left(Bucle) = Me.Controls(Bucle).left
        
        '
        RelW = Me.Height / Me_Height
        Me.Controls(Bucle).Height = RelW * CT_Height(Bucle)
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        Me.Controls(Bucle).top = RelW * ct_top(Bucle)
        ct_top(Bucle) = Me.Controls(Bucle).top
        On Error GoTo 0
        '
    Next Bucle
    Me_Width = Me.Width
    Me_Height = Me.Height
End Sub
Private Sub Borrar_Seccion(most As Long)
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Dim Reg2 As New RecordNet

    ' *****************
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & most)
    With Registro
        If Not .EOF Then
            MsgBox CargaCadena(1017)
            Exit Sub
        End If
    End With
    Set Registro = Base.OpenRecordset("select * from equipos where borrado=false and codi_ident=" & most)
    With Registro
        If Not .EOF Then
            MsgBox CargaCadena(1018)
            Exit Sub
        End If
    End With
    ' ********************
    Set Registro = Base.OpenRecordset("select * from seccion where borrado=false and codi_ident=" & most)
    With Registro
        If Not .EOF Then
            Reg2.OpenRecordset _
            ("select count(secc_maqui) from seccion where borrado=false and secc_maqui=" & .Fields("secc_Maqui"))
            If Reg2.Fields(0) > 1 Then
                Arbol.Nodes.Remove Arbol.SelectedItem.Index
                Delete_Record Registro
            Else
                Arbol.Nodes.Remove Arbol.SelectedItem.Index
                Edit_Record Registro
                .Fields("borrado") = True
                .Update
            End If
            .Close
            
            LblInfo2.Caption = CargaCadena(528)  '"Dato Eliminado.Envíe Modificaciones"
        End If
    End With
    '1.7.2
    Registro.Close
    Base.Close
    Set Reg2 = Nothing
    Set Registro = Nothing
    Set Base = Nothing
    'Workspaces(0).close

End Sub


Private Sub Arbol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then Exit Sub
    If Button = 2 Then
        If nodo = "sistema" Then
            PopupMenu mnuContextualSistema, 0, x + 200, y + 200
            Exit Sub
        End If
        If left(nodo, 1) = "S" Then PopupMenu mnuContextualSeccion, 0, x + 200, y + 200
        If left(nodo, 1) = "M" Then PopupMenu mnuContextualBalanza, 0, x + 200, y + 200
    End If
End Sub



Private Sub Arbol_NodeClick(ByVal Node As MSComctlLib.Node)
    nodo = Node.Key
    Informar (Node.Key)
End Sub

Private Sub CmdSalir_Click()
    LockBase = False
    Unload FrmMostrador2
    Unload FrmBalanza
    Unload Me
End Sub


Private Sub Command1_Click(Index As Integer)
    Me.Enabled = False
    Select Case Index
        Case 0
            Aux_Ejecuta "rpmostrador"
        Case 1
            Aux_Ejecuta "rpequipos"
    End Select
    Me.Enabled = True
End Sub

Private Sub Form_Load()
   'Dim Img As ListImage
   CambiarIdioma
    Me_Init
   Me.Icon = Form2.Icon
   Image1.Picture = Form2.Imagen.Picture
   FrmSeccionyMaquina.mnuContextualBalanza.Visible = False
   FrmSeccionyMaquina.mnuContextualSeccion.Visible = False
   FrmSeccionyMaquina.mnuContextualSistema.Visible = False
   Refresca_Arbol
   Informar ("sistema")
End Sub
Public Sub Refresca_Arbol()
    
    Dim Registro As New RecordNet
    Dim RegMaquina As New RecordNet
    Dim strSeguridad As String
    strSeguridad = ""
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = "and codi_ident=" & UsuarioActual.Asociado
    End If
    Arbol.Nodes.Clear
    Arbol.ImageList = Imagenes
    Arbol.Nodes.Add , , "sistema", CargaCadena(430), "pc"
    
    Registro.OpenRecordset ("select codi_ident from seccion where borrado=false " & strSeguridad & " order by codi_ident")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Arbol.Nodes.Add 1, tvwChild, "S" & .Fields("codi_ident"), .Fields("codi_ident"), "seccion"
                
                '
                ' meter equipos
                '
                RegMaquina.OpenRecordset _
                ("select * from equipos where borrado=false and codi_ident=" & .Fields("codi_ident"))
                With RegMaquina
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            If .Fields("modelo") < 100 Then
                                If .Fields("modelo") = 1 Then
                                    Arbol.Nodes.Add "S" & .Fields("codi_ident"), tvwChild, "M" & .Fields("numero_eqp"), .Fields("numero_eqp"), "g386M"
                                Else
                                    Arbol.Nodes.Add "S" & .Fields("codi_ident"), tvwChild, "M" & .Fields("numero_eqp"), .Fields("numero_eqp"), "g386"
                                End If
                            Else
                                If .Fields("modelo") = 100 Then Arbol.Nodes.Add "S" & .Fields("codi_ident"), tvwChild, "M" & .Fields("numero_eqp"), .Fields("numero_eqp"), "v12"
                                If .Fields("modelo") = 200 Then Arbol.Nodes.Add "S" & .Fields("codi_ident"), tvwChild, "M" & .Fields("numero_eqp"), .Fields("numero_eqp"), "redmarte"
                                If .Fields("modelo") = 400 Then Arbol.Nodes.Add "S" & .Fields("codi_ident"), tvwChild, "M" & .Fields("numero_eqp"), .Fields("numero_eqp"), "iv4lcd"

                            End If
                            .Movenext
                        Loop
                    End If
                End With
                '
                '
                
            .Movenext
            
            Loop
        End If
    End With
    Arbol.Nodes(1).Expanded = True

End Sub
Public Sub Informar(nodo As String)
    Dim Registro As New RecordNet
    info.TexT = ""
    If nodo = "sistema" Then
        AgregarInfo CargaCadena(430)  '("SISTEMA")
        AgregarInfo ("")
        If UsaGamaAlta Then
            If tcpip Then
                AgregarInfo CargaCadena(581)  '("Comunicación con Euroscale por TCP/IP")
            Else
                AgregarInfo CargaCadena(582)  '("Comunicación con Euroscale por RS-232")
            End If
        End If
        If usaSC10 Then
            If Not TipoC9 Then
                If SC10Rs232 Then
                    AgregarInfo CargaCadena(583)  '("Comunicación con SC10 por RS-232")
                Else
                    AgregarInfo CargaCadena(584)  '("Comunicación con SC10 por Path de Red")
                End If
            Else
                AgregarInfo CargaCadena(674) & " C9/C10"
                AgregarInfo ""
                If var.C9Cod = 0 Then
                    Select Case ModoC9
                        Case 0
                            AgregarInfo "3 " & CargaCadena(721) & " + 1 " & CargaCadena(618)
                            AgregarInfo CargaCadena(329) & " " & CargaCadena(721) & vbCrLf & "1,2,3 (max 999 PLUs)"
                            AgregarInfo CargaCadena(32) & " " & CargaCadena(618) & vbCrLf & "6 (max 499 PLUs)"
                        Case 1
                            AgregarInfo "5 " & CargaCadena(721) & " + 1 " & CargaCadena(618)
                            AgregarInfo CargaCadena(329) & " " & CargaCadena(721) & vbCrLf & "1,2,3,4,5 (max 599 PLUs)"
                            AgregarInfo CargaCadena(32) & " " & CargaCadena(618) & vbCrLf & "6 (max 499 PLUs)"
                        Case 2
                            AgregarInfo "7 " & CargaCadena(721) & " + 1 " & CargaCadena(618)
                            AgregarInfo CargaCadena(329) & " " & CargaCadena(721) & vbCrLf & "1,2,3,4,5,6,7 (max 399 PLUs)"
                            AgregarInfo CargaCadena(329) & " " & CargaCadena(618) & vbCrLf & "8 (max 499 PLUs)"
                    End Select
                Else
                    AgregarInfo "1 sección con 7000 artículos"
                End If
            End If
        End If
        If UsaECO1 Then
            If TipoV8 Then
                AgregarInfo "V8 - RS-232"
            Else
                AgregarInfo "Marte IC-ILC - RS-232"
            End If
        End If
        If UsaECO2 Then
            AgregarInfo "Marte IC-ILC" ' - RS-232" 2.0.a
        End If
        If UsaEcoPrint Then
            Select Case TipoEcoPrint
                Case 0
                    AgregarInfo "IV4 STD - RS-232"
                Case 1
                    AgregarInfo "IV4 CE - RS-232"
                Case 2
                    AgregarInfo "IV4 TQ - RS-232"
                Case 3
                    AgregarInfo "IV4 LCD - RS-232 - COM " & PuertoEcoPrint
            End Select
        End If
        Exit Sub
    End If
    '**************
    ' Info de secciones
    '**************
    If left(nodo, 1) = "S" Then
        nodo = Right(nodo, Len(nodo) - 1)
        Registro.OpenRecordset ("select * from seccion where borrado=false and codi_ident=" & nodo)
        With Registro
        If Not .EOF Then
            AgregarInfo CargaCadena(302)  '("MOSTRADOR")
            AgregarInfo .Fields("nombre")
            AgregarInfo (Chr(34) & .Fields("descripcio") & Chr(34))
            'AgregarInfo ("Mostrador número " & .Fields("codi_ident)
            'AgregarInfo ("Sección Máquina " & .Fields("secc_maqui)
            AgregarInfo (CargaCadena(594) & " " & .Fields("codi_ident"))
            AgregarInfo (CargaCadena(54) & " " & .Fields("secc_Maqui"))
            Select Case .Fields("enviardatos")
                Case "GA"
                    AgregarInfo CargaCadena(587)  '("Envío de artículos sólo a Euroscale")
                    If .Fields("multiple") = 1 Then
                        AgregarInfo CargaCadena(585)  '("Algunos artículos marcados para enviar a SC10")
                    End If
                    If MultiEuroscale Then AgregarInfo ("IP : " & .Fields("sec_ip"))
                    
                Case "SC10"
                    AgregarInfo CargaCadena(588)  '("Envío de artículos sólo a Concentrador")
                    If .Fields("multiple") = 1 Then
                        AgregarInfo CargaCadena(586)  '("Algunos artículos marcados para enviar a Euroscale")
                    End If
                Case "ECON"
                    If TipoV8 Then
                        AgregarInfo CargaCadena(1110)
                    Else
                        AgregarInfo CargaCadena(1112)
                    End If
                Case "ECOL"
                    AgregarInfo CargaCadena(1111)
                    If Not IsNull(.Fields("multiple")) And (.Fields("multiple") <> 0) Then
                        AgregarInfo ("IP : " & .Fields("sec_ip") & ":" & CStr(.Fields("multiple")))
                    Else
                        AgregarInfo ("RS-232" & " COM" & CStr(.Fields("com")))
                    End If
                Case "ECOP"
                    AgregarInfo CargaCadena(1120)
                Case "TODOS"
                    AgregarInfo CargaCadena(589)  '("Envío de artículos a Ambos Sistemas")
            
            End Select
            
        End If
        End With
    End If
    '*************
    ' Info de máquinas
    '*************
    If left(nodo, 1) = "M" Then
        nodo = Right(nodo, Len(nodo) - 1)
        Registro.OpenRecordset ("select * from equipos where borrado=false and numero_eqp=" & _
        nodo)
        With Registro
        If Not .EOF Then
            
            If .Fields("modelo") = 100 Then
                AgregarInfo CargaCadena(591)  '("MAQUINA V-12")
            Else
                Select Case .Fields("modelo")
                    Case 0
                        AgregarInfo CargaCadena(590)  '("MAQUINA EUROSCALE")
                    Case 1
                        AgregarInfo CargaCadena(590) & " (SRV)" '("MAQUINA EUROSCALE")
                    Case 200
                        
                        AgregarInfo "Marte IC-ILC / V-8"


                    Case 400
                        Select Case TipoEcoPrint
                            Case 0
                                AgregarInfo "IV4 STD"
                            Case 1
                                AgregarInfo "IV4 CE"
                            Case 2
                                AgregarInfo "IV4 TQ"
                            Case 3
                                AgregarInfo "IV4 LCD"
                        End Select
                End Select
            End If
            AgregarInfo ("")
             AgregarInfo (Chr(34) & .Fields("descripcio") & Chr(34))
            'AgregarInfo ("Número de máquina : " & .Fields("numero_eqp)
            'AgregarInfo ("Asociada al mostrador : " & .Fields("codi_ident)
            'AgregarInfo ("Sección Máquina : " & .Fields("secc_maqui)
            AgregarInfo (CargaCadena(568) & " " & .Fields("numero_eqp"))
            AgregarInfo (CargaCadena(593) & " " & .Fields("codi_ident"))
            AgregarInfo (CargaCadena(54) & " " & .Fields("secc_Maqui"))
            If .Fields("modelo") < 100 And tcpip Then
                'AgregarInfo ("Dirección IP : " & .Fields("prog_eqp)
                AgregarInfo (CargaCadena(569) & " " & .Fields("prog_eqp"))
            End If
            If .Fields("modelo") = 1 Then AgregarInfo CargaCadena(592)  '("Balanza Servidora")
               
        End If
        End With
    End If
   
End Sub
Private Sub AgregarInfo(cadena As String)
    info.TexT = info.TexT & cadena & Chr(13) & Chr(10)
End Sub

Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Dir(App.Path & "\servidoratouch.ini") <> "" Then
        Kill App.Path & "\servidoratouch.ini"
    End If
    frmEpelsa.Enabled = True
End Sub





Private Sub mnuCreaBalanza_Click()
    If Right(Arbol.SelectedItem.Key, Len(Arbol.SelectedItem.Key) - 1) <> "0" Then
        Load FrmBalanza
        Arbol.SelectedItem.Expanded = True
        FrmBalanza.SeccionAsignada = Right(Arbol.SelectedItem.Key, Len(Arbol.SelectedItem.Key) - 1)
        FrmBalanza.Show
        Me.Enabled = False
        Do While FrmBalanza.Visible
            Do_Events
        Loop
        Me.Enabled = True
    End If
End Sub

Private Sub mnuCreaSeccion_Click()
    Load FrmMostrador2
    FrmMostrador2.CmbMostrador.Enabled = True
    Me.Enabled = False
    FrmMostrador2.Show
    Do While FrmMostrador2.Visible
        Do_Events
    Loop
    Me.Enabled = True
End Sub

Private Sub mnuEliminaBalanza_Click()
    Dim respuesta As String
    Dim Miequipo As DB_Balanzas
    
    respuesta = MsgBox(CargaCadena(914), vbYesNo)
    If respuesta = vbNo Then Exit Sub
    Load FrmBalanza
    FrmBalanza.CmbEquipo.TexT = Arbol.SelectedItem.TexT
    ' CmdBorrar_Click
    Miequipo.NBalanza = FrmBalanza.CmbEquipo.TexT
    If Baja_Balanza(Miequipo) = 0 Then
        FrmBalanza.LblInfo2.Caption = CargaCadena(528)  '"Dato Eliminado.Envíe Modificaciones"
        FrmBalanza.CmbEquipo.TexT = ""
        FrmBalanza.refresca_equipos
        FrmBalanza.Refresca_Datos
        FrmSeccionyMaquina.Refresca_Arbol
    End If
    
    Unload FrmBalanza
    info.TexT = ""
    LblInfo2.Caption = CargaCadena(528)  '"Datos Eliminados.Envíe Modificaciones"
End Sub

Private Sub mnuEliminaMostrador_Click()
    Dim respuesta As String
    Dim Registro As New RecordNet
    Dim MiMostrador As DB_mostrador
    respuesta = MsgBox(CargaCadena(915), vbYesNo)
    If respuesta = vbNo Then Exit Sub
    MiMostrador.NMostrador = Arbol.SelectedItem.TexT
    Select Case Baja_Mostrador(MiMostrador)
        Case 0
            LblInfo2.Caption = CargaCadena(528)  '"Datos Eliminados.Envíe Modificaciones"
            FrmSeccionyMaquina.Refresca_Arbol
        Case 10
            MsgBox CargaCadena(404)
        Case 11
            MsgBox CargaCadena(403)
    End Select
    info.TexT = ""
End Sub

Private Sub mnuModificaBalanza_Click()
    Load FrmBalanza
  
    FrmBalanza.SeccionAsignada = Right(Arbol.SelectedItem.Parent.Key, Len(Arbol.SelectedItem.Parent.Key) - 1)
    FrmBalanza.CmbTipo.Enabled = False
    FrmBalanza.CmbEquipo.Enabled = False
    FrmBalanza.CmbEquipo.TexT = Right(nodo, Len(nodo) - 1)
    FrmBalanza.CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
    FrmBalanza.LblInfo.BackColor = vbYellow
    FrmBalanza.LblInfo.Caption = CargaCadena(525)  '"Modificación"
    FrmBalanza.CmdBorrar.Enabled = False
    FrmBalanza.Refresca_Datos
    Me.Enabled = False
    FrmBalanza.Show
    Do While FrmBalanza.Visible
        Do_Events
    Loop
    Me.Enabled = True
End Sub

Private Sub mnuModificaMostrador_Click()
    
    Dim nF, ncont As Integer
    Dim Registro As New RecordNet

    Registro.OpenRecordset _
    ("select * from seccion where borrado=false and codi_ident=" _
    & Right(Arbol.SelectedItem.Key, Len(Arbol.SelectedItem.Key) - 1))
  
    FrmMostrador2.CmbMostrador.TexT = Format(Right(nodo, Len(nodo) - 1), "000")
    
    If Dir(App.Path & "\plrm.txt") <> "" Then
        nContSecPlrm = 0
        nF = FreeFile()
        Open App.Path & "\plrm.txt" For Input As #nF
        Do While Not EOF(nF)
            nContSecPlrm = nContSecPlrm + 1
            Line Input #nF, sPathSecPlrm(nContSecPlrm)
            sPathSecPlrm(nContSecPlrm) = Trim(sPathSecPlrm(nContSecPlrm))
            If Right(sPathSecPlrm(nContSecPlrm), 1) <> "\" Then sPathSecPlrm(nContSecPlrm) = sPathSecPlrm(nContSecPlrm) & "\"
            nSecPLRM(nContSecPlrm) = Val(Mid(sPathSecPlrm(nContSecPlrm), 1, 2))
            sPathSecPlrm(nContSecPlrm) = Mid(sPathSecPlrm(nContSecPlrm), 3)
        Loop
        Close #nF
        For ncont = 1 To nContSecPlrm
            If Val(FrmMostrador2.CmbMostrador.TexT) = Val(nSecPLRM(ncont)) Then
                FrmMostrador2.TextPathSec.TexT = sPathSecPlrm(ncont)
            End If
        Next ncont
    End If
    
    FrmMostrador2.CmbMostrador.Enabled = False
    FrmMostrador2.TxtMaquina.Enabled = False
    FrmMostrador2.TxtMaquina.TexT = Registro.Fields("secc_Maqui")
    FrmMostrador2.Text1.TexT = Registro.Fields("nombre")
    FrmMostrador2.Text2.TexT = Registro.Fields("descripcio")
    If Not IsNull(Registro.Fields("SEC_IP")) Then
        FrmMostrador2.tXTip.TexT = Registro.Fields("SEC_IP")
    End If
    If Not (IsNull(Registro.Fields("multiple"))) And (Registro.Fields("multiple") <> 0) Then
        FrmMostrador2.TextPeco.TexT = CStr(Registro.Fields("multiple"))
    End If
    FrmMostrador2.Command1.Enabled = False
    Select Case Registro.Fields("enviardatos")
        Case "GA"
            FrmMostrador2.CmbEnviar.TexT = FrmMostrador2.CmbEnviar.List(0)
        Case "SC10"
            FrmMostrador2.CmbEnviar.TexT = FrmMostrador2.CmbEnviar.List(1)
        Case "ECON"
            FrmMostrador2.CmbEnviar.TexT = FrmMostrador2.CmbEnviar.List(2)
        Case "ECOL"
            FrmMostrador2.CmbEnviar.TexT = FrmMostrador2.CmbEnviar.List(3)
        Case "TODOS"
            FrmMostrador2.CmbEnviar.TexT = FrmMostrador2.CmbEnviar.List(4)
    End Select
    If Registro.Fields("enviardatos") = "ECON" Or Registro.Fields("enviardatos") = "ECOL" Then
        FrmMostrador2.CmbCOM.Enabled = True
        FrmMostrador2.Label2(3).Enabled = True
    Else
        FrmMostrador2.CmbCOM.Enabled = False
        FrmMostrador2.Label2(3).Enabled = False
    End If
    'If IsNumeric(Registro.Fields("COM")) Then
    '    On Error Resume Next
    '    '2.0.36
    '    'If AhorraMas And usaSC10 And Registro.Fields("com") = 255 Then
    '    '    FrmMostrador2.ChkAhoSS.Value = vbChecked
    '    'Else
    '    '    FrmMostrador2.ChkAhoSS.Value = vbUnchecked
    '    'End If
    '    '''''''''''''''''''''''''''''''''
    '    FrmMostrador2.CmbCOM.ListIndex = Registro.Fields("COM")
    '    On Error GoTo 0
    'Else
        FrmMostrador2.CmbCOM.ListIndex = 0
    'End If
    FrmMostrador2.LblInfo.Caption = CargaCadena(525)  '"Modificación"
    FrmMostrador2.LblInfo.BackColor = vbYellow
    FrmMostrador2.CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"

    Me.Enabled = False
    FrmMostrador2.Show
    Do While FrmMostrador2.Visible
        Do_Events
    Loop
    Me.Enabled = True
    
    '2.0.a
    Call Informar(nodo)
    Me.Refresh
    ''''''
End Sub
Private Sub CambiarIdioma()
    FrmSeccionyMaquina.Caption = CargaCadena(580)
    Command1(0).Caption = CargaCadena(578)
    Command1(1).Caption = CargaCadena(579)
    CmdSalir.Caption = CargaCadena(7)
    mnuModificaBalanza.Caption = CargaCadena(595)
    mnuEliminaBalanza.Caption = CargaCadena(596)
    mnuModificaMostrador.Caption = CargaCadena(597)
    mnuEliminaMostrador.Caption = CargaCadena(598)
    mnuCreaBalanza.Caption = CargaCadena(599)
    mnuCreaSeccion.Caption = CargaCadena(600)
End Sub

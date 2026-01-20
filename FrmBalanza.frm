VERSION 5.00
Begin VB.Form FrmBalanza 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E"
   ClientHeight    =   3945
   ClientLeft      =   3000
   ClientTop       =   2445
   ClientWidth     =   4770
   Icon            =   "FrmBalanza.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Checkss165 
      Caption         =   "SS 165."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   2052
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3720
      TabIndex        =   13
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   12
      Top             =   3480
      Width           =   972
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   972
   End
   Begin VB.ComboBox CmbEquipo 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   2532
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      ItemData        =   "FrmBalanza.frx":000C
      Left            =   2640
      List            =   "FrmBalanza.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2052
   End
   Begin VB.CheckBox ChkServidora 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   4572
   End
   Begin VB.Label LblMostrador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   2772
   End
   Begin VB.Label LblInfo2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H80000009&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label LblInfoIP 
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label LblIP 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2172
   End
End
Attribute VB_Name = "FrmBalanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LaSeccion As Long
Private HayServidora As Boolean
Private QL As Boolean
Private NoMostrarIP As Boolean
Public Property Let SeccionAsignada(ByVal Mostrador As Long)
    
    Dim Registro As New RecordNet
    Dim MiSec As Integer
    Dim MIERROR As Boolean
    '***********************************
    ' Control para no dejar introducir balanzas
    ' V-12 en mostradores incorrectos para
    ' C9/C10
    '***********************************
    LaSeccion = Mostrador
    LblMostrador.Caption = CargaCadena(302) & " " & LaSeccion
    CmdAceptar.Enabled = True
    If usaSC10 And TipoC9 And CmbTipo.ListIndex = 1 Then
        MiSec = 0
        
        Registro.OpenRecordset ("SELECT secc_maqui FROM SECCION where codi_ident=" & Mostrador & " and borrado=false")
        With Registro
            If Not .EOF Then
                MiSec = .Fields("secc_Maqui")
            End If
        End With
       
        MIERROR = False
        Select Case ModoC9
            Case 0
                If (MiSec < 1 Or MiSec > 3) And MiSec <> 6 Then MIERROR = True
            Case 1
                If (MiSec < 1 Or MiSec > 5) And MiSec <> 6 Then MIERROR = True
            Case 2
                If (MiSec < 1 Or MiSec > 7) And MiSec <> 8 Then MIERROR = True
        End Select
        If MIERROR Then
            If UsaGamaAlta Then
                CmbTipo.Enabled = False
            Else
                CmdAceptar.Enabled = False
            End If
        End If
    End If
End Property

Public Sub refresca_equipos()
    
    Dim Registro As New RecordNet
    '********************************
    ' sitúa en el combo los
    ' equipos correspondientes
    ' a la sección elegida
    '********************************
    CmbEquipo.Clear

    Registro.OpenRecordset _
    ("select numero_eqp,borrado,codi_ident from equipos where borrado=false and codi_ident=" & LaSeccion & " order by numero_eqp")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                CmbEquipo.AddItem .Fields("numero_eqp")
                .Movenext
            Loop
        End If
    End With
  
End Sub
Public Sub Refresca_Datos()
    
    Dim Registro As New RecordNet
    '*************************************
    ' Sitúa los datos que correspondan con
    ' la balanza seleccionada
    '*************************************
    '**********************
    ' 1º limpia todos
    ' los datos
    tXTip.TexT = "___.___.___.___"
    If NoMostrarIP Then tXTip.TexT = "127.0.0.1"
    ChkServidora.Value = vbUnchecked
    TxtDescripcion.TexT = ""
    LblInfo.BackColor = vbWhite
    LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
    CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
    If CmbTipo.TexT <> CmbTipo.List(0) Then
        '1.9.0
        'Checkss165.Enabled = False
        ''''''
        ChkServidora.Enabled = False
    Else
        
        '1.9.0
        'Checkss165.Enabled = True
        ''''''
        
        If Not Comprobar_Servidora Then
            ChkServidora.Enabled = True
        Else
            ChkServidora.Enabled = False
        End If
    End If
    '**********************
    ' si existe el equipo
    ' sitúa sus datos
    '**********************
    If IsNumeric(CmbEquipo.TexT) Then
       
        Registro.OpenRecordset _
        ("select * from equipos where borrado=false and numero_eqp=" & Val(CmbEquipo.TexT))
        With Registro
            If Not .EOF Then
                If LaSeccion = .Fields("codi_ident") Then
                    If .Fields("modelo") = 200 Then CmbTipo.TexT = CmbTipo.List(2)
                    If .Fields("modelo") = 400 Then CmbTipo.TexT = CmbTipo.List(3)
                    If .Fields("modelo") = 100 Then CmbTipo.TexT = CmbTipo.List(1)
                    If .Fields("modelo") < 100 Then CmbTipo.TexT = CmbTipo.List(0)
                    
                    If .Fields("modelo") >= 100 Then
                        tXTip.TexT = "___.___.___.___"
                        ChkServidora.Value = vbUnchecked
                        LblIP.Enabled = False
                        ChkServidora.Enabled = False
                        tXTip.Enabled = False
                    End If
                    If .Fields("modelo") < 100 Then
                        tXTip.TexT = .Fields("prog_eqp")
                        If .Fields("modelo") = 0 Then ChkServidora.Value = vbUnchecked
                        If .Fields("modelo") = 1 Then ChkServidora.Value = vbChecked
                        'If .Fields("ss165") Then
                        '    Checkss165.Value = vbChecked
                        'Else
                        '    Checkss165.Value = vbUnchecked
                        'End If
                        
                        If Not NoMostrarIP Then ChkServidora.Enabled = True
                        If tcpip And (Not NoMostrarIP) Then
                            LblIP.Enabled = True
                            tXTip.Enabled = True
                        End If
                    End If
                    TxtDescripcion.TexT = .Fields("descripcio")
                    LblInfo.BackColor = vbYellow
                    LblInfo.Caption = CargaCadena(525)  '"Modificación"
                    CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
                Else
                    'MsgBox "La balanza " & .Fields(numero_eqp & " ya existe, y pertenece al mostrador " & .Fields(codi_ident
                    MsgBox CargaCadena(573) & " " & .Fields("numero_eqp") & " " & CargaCadena(574) & " " & .Fields("codi_ident")
                    CmbEquipo.TexT = ""
                End If
            End If
            
        End With
    
        If CmbEquipo.Enabled = False _
        And Comprobar_Servidora And _
        ChkServidora.Value = vbUnchecked Then
            ChkServidora.Enabled = False
        Else
            If CmbTipo.TexT = CmbTipo.List(0) And Not Comprobar_Servidora Then ChkServidora.Enabled = True
        End If
    End If
End Sub
Private Function Comprobar_Servidora() As Boolean
    
    Dim Registro As New RecordNet
    Dim Conforme As Boolean
    '************************************
    ' Comprueba si ya existe una balanza
    ' servidora Euroscale, para no dejar
    ' introducir otra
    '************************************
    Conforme = False
   
    Registro.OpenRecordset ("select numero_eqp from equipos where borrado=false and modelo=1")
    With Registro
        If Not .EOF Then
            'LblInfoIP.Caption = "La balanza servidora es la Nº " & .Fields(numero_eqp
            LblInfoIP.Caption = CargaCadena(570) & .Fields("numero_eqp")
            Conforme = True
        Else
            LblInfoIP.Caption = ""
            Conforme = False
        End If
    End With

    Comprobar_Servidora = Conforme
End Function

Private Sub ChkServidora_Click()
    If Dir(App.Path & "\servidoratouch.ini") <> "" Then
        Kill App.Path & "\servidoratouch.ini"
    End If
End Sub

Private Sub CmbEquipo_Click()
    Refresca_Datos
End Sub
Private Sub CmbEquipo_KeyPress(KeyAscii As Integer)
    CmbEquipo.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If Len(CmbEquipo.TexT) >= 2 Then CmbEquipo.Locked = True
    If KeyAscii < 48 Or KeyAscii > 57 Then CmbEquipo.Locked = True
End Sub

Private Sub CmbEquipo_LostFocus()
    Refresca_Datos
End Sub

Private Sub CmbTipo_Click()
    Dim TipoIncorrecto As Boolean
    '******************************
    ' corrige al usuario si elige un tipo
    ' de equipo incorrecto
    '*******************************
    If QL Then Exit Sub
    Select Case CmbTipo.ListIndex
        Case 0
            If Not UsaGamaAlta Then TipoIncorrecto = True
        Case 1
            If Not usaSC10 Then TipoIncorrecto = True
        Case 2
            If (Not UsaECO1) And (Not UsaECO2) Then TipoIncorrecto = True
        Case 3
            If Not UsaEcoPrint Then TipoIncorrecto = True
    End Select
    If TipoIncorrecto Then
        QL = True
        MsgBox CargaCadena(1179), vbCritical
        If UsaEcoPrint Then CmbTipo.ListIndex = 3
        If (UsaECO2 Or UsaECO1) Then CmbTipo.ListIndex = 2
        If usaSC10 Then CmbTipo.ListIndex = 1
        If UsaGamaAlta Then CmbTipo.ListIndex = 0
        QL = False
    End If
    If CmbTipo.ListIndex <> 0 Then
        LblIP.Enabled = False
        tXTip.Enabled = False
        ChkServidora.Enabled = False
        LblInfoIP.Visible = False
        '1.9.0
        'Checkss165.Enabled = False
        ''''''
    Else
        If Not NoMostrarIP Then
            LblIP.Enabled = True
            tXTip.Enabled = True
        End If
        ChkServidora.Enabled = True
        '1.9.0
        'Checkss165.Enabled = True
        ''''''
        LblInfoIP.Visible = True
    End If
    SeccionAsignada = LaSeccion
End Sub

Private Sub CmdAceptar_Click()
    Dim Miequipo As DB_Balanzas
    Dim MiResultado As Integer
    '******************************
    ' Introducir los datos indicados
    ' en la base (alta o modificación)
    '******************************
    If CmbEquipo.TexT = "" Then
        MsgBox CargaCadena(575), vbCritical  '"Introduzca un número de equipo"
        CmbEquipo.SetFocus
        Exit Sub
    End If
    Miequipo.NBalanza = CmbEquipo.TexT
    Miequipo.NMostrador = LaSeccion
    Miequipo.DireccionIP = tXTip.TexT
    Miequipo.Tipo = CmbTipo.ListIndex * 100
    If CmbTipo.TexT = CmbTipo.List(0) Then
        'If Checkss165.Value = vbChecked Then
        '    Miequipo.lss165 = True
        'Else
            Miequipo.lss165 = False
        'End If
        If (ChkServidora.Value = vbChecked) And (Not MultiEuroscale) Then
            Miequipo.Tipo = Miequipo.Tipo + 1
        End If
    Else
        Miequipo.lss165 = False
    End If
    If Miequipo.Tipo = 300 Then Miequipo.Tipo = 400
    Miequipo.Descripcion = left(TxtDescripcion.TexT, 21)
    MiResultado = Alta_Balanza(Miequipo)
    If MiResultado = 60 Then
        MsgBox CargaCadena(1202), vbCritical
        Exit Sub
    End If
    If MiResultado = 12 Then
        MsgBox CargaCadena(576), vbCritical  '"El número de Equipo debe estar comprendido entre 1 y 60"
        CmbEquipo.SetFocus
        Exit Sub
    End If
    If MiResultado = 52 Then
        MsgBox CargaCadena(1118), vbCritical
        Exit Sub
    End If
    If MiResultado = 53 Then
        MsgBox CargaCadena(1193), vbCritical
        Exit Sub
    End If
    If MiResultado = 13 Then
        MsgBox CargaCadena(577), vbCritical  '"Falta la dirección IP"
        tXTip.SetFocus
        Exit Sub
    End If
    If MiResultado = 1 Then
        
        LblInfo2.Caption = CargaCadena(439)  '"Dato Modificado.Envíe Modificaciones"
        On Error Resume Next
        FrmSeccionyMaquina.Arbol.Nodes.Remove ("M" & CmbEquipo.TexT)
        On Error GoTo 0
        If CmbEquipo.Enabled = True Then CmbEquipo.SetFocus
    End If
    If MiResultado = 0 Then
        
        LblInfo2.Caption = CargaCadena(527)  '"Dato Añadido.Envíe Modificaciones"
        
        If CmbEquipo.Enabled = True Then CmbEquipo.SetFocus
    End If
    If MiResultado < 2 Then
        If CmbTipo.TexT = CmbTipo.List(0) Then
            If ChkServidora.Value = vbChecked Then
                FrmSeccionyMaquina.Arbol.Nodes.Add "S" & LaSeccion, tvwChild, "M" & CmbEquipo.TexT, CmbEquipo.TexT, "g386M"
            Else
                FrmSeccionyMaquina.Arbol.Nodes.Add "S" & LaSeccion, tvwChild, "M" & CmbEquipo.TexT, CmbEquipo.TexT, "g386"
            End If
        Else
            If CmbTipo.TexT = CmbTipo.List(1) Then FrmSeccionyMaquina.Arbol.Nodes.Add "S" & LaSeccion, tvwChild, "M" & CmbEquipo.TexT, CmbEquipo.TexT, "v12"
            If CmbTipo.TexT = CmbTipo.List(2) Then FrmSeccionyMaquina.Arbol.Nodes.Add "S" & LaSeccion, tvwChild, "M" & CmbEquipo.TexT, CmbEquipo.TexT, "redmarte"
            If CmbTipo.TexT = CmbTipo.List(3) Then FrmSeccionyMaquina.Arbol.Nodes.Add "S" & LaSeccion, tvwChild, "M" & CmbEquipo.TexT, CmbEquipo.TexT, "iv4lcd"
        End If
    If CmbEquipo.Enabled = False Then FrmSeccionyMaquina.Informar (FrmSeccionyMaquina.Arbol.SelectedItem.Key)
    CmbEquipo.TexT = ""
    refresca_equipos
    Refresca_Datos
    If CmbEquipo.Enabled = False Then
        FrmSeccionyMaquina.LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados.Envíe Modificaciones"
        Unload Me
    End If
    End If
    
End Sub


'1.0.4
'Public Sub CmdBorrar_Click()
    'Dim Miequipo As DB_Balanzas
    ''******************************
    '' Orden de borrado
    ''******************************
    'If IsNumeric(CmbEquipo.TexT) Then
    '    Miequipo.NBalanza = CmbEquipo.TexT
    '    If Baja_Balanza(Miequipo) = 0 Then
    '        LblInfo2.Caption = CargaCadena(528)  '"Dato Eliminado.Envíe Modificaciones"
    '        CmbEquipo.TexT = ""
    '        refresca_equipos
    '        Refresca_Datos
    '        FrmSeccionyMaquina.Refresca_Arbol
    '    End If
    'End If
'End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim valorip As String
    NoMostrarIP = False
    valorip = Check_Dns_File()
    If valorip <> "" Then
        NoMostrarIP = True
    End If
    
    CambiarIdioma
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    If UsaEcoPrint Then CmbTipo.TexT = CmbTipo.List(3)
    If UsaECO2 Then CmbTipo.TexT = CmbTipo.List(2)
    If UsaECO1 Then CmbTipo.TexT = CmbTipo.List(2)
    If usaSC10 Then CmbTipo.TexT = CmbTipo.List(1)
    If UsaGamaAlta Then CmbTipo.TexT = CmbTipo.List(0)
    If NumeroSistemas > 1 Then
        CmbTipo.Enabled = True
    End If
    If Comprobar_Servidora Then
        ChkServidora.Enabled = False
        HayServidora = True
    Else
        ChkServidora.Enabled = True
        HayServidora = False
    End If
    If (tcpip = False) Or (NoMostrarIP) Then
        tXTip.Enabled = False
        LblIP.Enabled = False
    End If
    ChkServidora.Visible = True
    If MultiEuroscale Or (NoMostrarIP) Then ChkServidora.Visible = False
    refresca_equipos
    Refresca_Datos
End Sub
Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    TxtDescripcion.Locked = False
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If Len(TxtDescripcion.TexT) = 21 Then TxtDescripcion.Locked = True
End Sub

Private Sub TxtIP_GotFocus()
    If tXTip.TexT = "___.___.___.___" Then tXTip.TexT = ""
End Sub

Private Sub tXTip_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case Asc(".")
        Case vbKeyBack
        Case vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
        
End Sub

Private Sub txtip_LostFocus()
    Dim sData As String
    sData = tXTip.TexT
    If tXTip.TexT <> "___.___.___.___" Then
        If Not GeneraTextoIP(sData) Then
            LblInfo2.Caption = CargaCadena(572)  '"Formato IP incorrecto"
            tXTip = "___.___.___.___"
        Else
            tXTip.TexT = sData
        End If
    End If
End Sub
Private Sub CambiarIdioma()
    FrmBalanza.Caption = CargaCadena(571)
    LblInfo.Caption = CargaCadena(524)
    Label1.Caption = CargaCadena(567)
    Label2.Caption = CargaCadena(568)
    LblIP.Caption = CargaCadena(569)
    ChkServidora.Caption = CargaCadena(492)
    Label3.Caption = CargaCadena(57)
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
End Sub

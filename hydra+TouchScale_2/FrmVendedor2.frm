VERSION 5.00
Begin VB.Form FrmVendedor2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Vendedores"
   ClientHeight    =   3645
   ClientLeft      =   4290
   ClientTop       =   405
   ClientWidth     =   6210
   Icon            =   "FrmVendedor2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Text1 
      Height          =   315
      ItemData        =   "FrmVendedor2.frx":000C
      Left            =   2160
      List            =   "FrmVendedor2.frx":000E
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1212
   End
   Begin VB.TextBox CmbNombre 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   3972
   End
   Begin VB.CommandButton CmdListado 
      Caption         =   "Listado"
      Height          =   372
      Left            =   4920
      TabIndex        =   7
      Top             =   2400
      Width           =   1212
   End
   Begin VB.ComboBox CmbTecla 
      Height          =   315
      ItemData        =   "FrmVendedor2.frx":0010
      Left            =   2160
      List            =   "FrmVendedor2.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   1092
   End
   Begin VB.ComboBox CmbMostrador 
      Height          =   315
      ItemData        =   "FrmVendedor2.frx":0014
      Left            =   2160
      List            =   "FrmVendedor2.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3972
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Width           =   1092
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   1092
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Añadir"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   1092
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Width           =   3972
   End
   Begin VB.Label Label2 
      Caption         =   "Identificación"
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
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Tecla de Vendedor"
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
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Mostrador"
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   252
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   2172
   End
End
Attribute VB_Name = "FrmVendedor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const BALSAM = False

Private Type TipoMostradores
    Mostrador() As Integer
    cantidad As Integer
End Type


Private Sub Refresca_Datos()
    
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim bucle As Integer
    Dim Buffer As String
    Dim Token As Long
    
    CmbMostrador.Enabled = True
    CmbNombre.TexT = ""
    LblInfo.BackColor = vbWhite
    LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
    CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
    If Text1.TexT = "" Then Exit Sub
 
    If CmbMostrador.TexT = "" Then
        If CmbMostrador.ListCount > 0 Then
            CmbMostrador.ListIndex = 0
        End If
    End If

    Registro.OpenRecordset _
    ("select * from vendedor where borrado=false and ident_vend=" & Text1.TexT & " and secc_maqui=" & CStr(Val(CmbMostrador.TexT)))
    With Registro
        If Not .EOF Then
            If .Fields("prog_vend") = "*" Or .Fields("prog_vend") = "-" Then
#If BALSAM = True Then
                If Dir(App.Path & "\" & "notecvend") = "" Then
                    CmbTecla.TexT = .Fields("tec_vend")
                End If
#Else
                If .Fields("tec_vend") > 0 And .Fields("tec_vend") < 31 Then
                    CmbTecla.TexT = .Fields("tec_vend")
                End If
#End If
            Else
                CmbMostrador.TexT = 0
            End If
            CmbNombre.TexT = .Fields("Nombre")
            LblInfo.BackColor = vbYellow
            LblInfo.Caption = CargaCadena(525)  '"Modificación"
            CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
        End If
    End With

End Sub
Private Sub refresca_Nombres()
    
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    
    Text1.Clear
'#If BALSAM = false Then
    Registro.OpenRecordset _
    ("select distinct(ident_vend) from vendedor where borrado=false order by ident_vend")
'#Else
'    Registro.OpenRecordset _
'    ("select distinct(ident_vend) from vendedor where borrado=false order by ident_vend")
'#End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Text1.AddItem .Fields("ident_vend")
'#If BALSAM = false Then
'                If .Fields("tec_vend") <> 0 Then
'                    CmbTecla.AddItem .Fields("TEC_VEND")
'                Else
'                    CmbTecla.AddItem .Fields("ident_vend")
'                End If
'#End If
               .Movenext
            Loop
        End If
    End With
  
End Sub

'Private Sub refresca_teclas()
'
'    Dim Registro As New RecordNet
'    Dim Reg2 As New RecordNet
'    CmbTeclas.Clear
'
'    Registro.OpenRecordset _
'    ("select distinct(ident_vend) from vendedor where borrado=false order by ident_vend")
'    With Registro
'        If Not .EOF Then
'            .MoveFirst
'            Do Until .EOF
'                CmbTeclas.AddItem .Fields("tec_vend")
'               .Movenext
'            Loop
'        End If
'    End With
'
'End Sub

Private Sub Refresca_Mostrador()
    
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim Buffer As String
    CmbMostrador.Clear
    Registro.OpenRecordset _
    ("select distinct secc_maqui,borrado from seccion where (borrado=false and secc_maqui<>0) order by secc_maqui")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                Reg2.OpenRecordset _
                ("select codi_ident,secc_maqui from seccion where borrado=false and secc_maqui=" & .Fields("secc_Maqui") & " order by codi_ident")
                With Reg2
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            Buffer = Buffer & .Fields("codi_ident") & " , "
                            .Movenext
                        Loop
                        Buffer = left(Buffer, Len(Buffer) - 3)
                        CmbMostrador.AddItem Buffer
                    End If
                End With
               .Movenext
            Loop
        End If
    End With

End Sub

Private Sub CmbMostrador_click()
    Call Refresca_Datos
End Sub

Private Sub CmbNombre_KeyPress(KeyAscii As Integer)
    CmbNombre.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If Len(CmbNombre.TexT) >= 30 Then CmbNombre.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LockBase = False
    frmEpelsa.Enabled = True
End Sub

Private Sub Text1_LostFocus()
    Refresca_Datos
End Sub

Private Sub Cmdaceptar_Click()
    
    Dim Registro As New RecordNet
    Dim bucle As Long
    Dim Mostradores As TipoMostradores
    Dim MiVendedor As DB_Vendedor
    Dim Retorno As Integer
    ' *********************
    ' Condiciones de error
    If Text1.TexT = "" Then Exit Sub
    If CmbMostrador.TexT = "" Or (CmbMostrador.TexT = "0") Then ' And OptAlta(0).Value = True) Then
        MsgBox CargaCadena(559)  '"Introduzca el número de mostrador"
        CmbMostrador.SetFocus
        Exit Sub
    End If
#If BALSAM = True Then
    If Dir(App.Path & "\" & "notecvend") = "" Then
        If CmbTecla.TexT = "" Or (CmbTecla.TexT = "0") Then ' And OptAlta(0).Value = True) Then
            MsgBox CargaCadena(1377) 'CargaCadena(645)  '"Introduzca Tecla"
            CmbTecla.SetFocus
            Exit Sub
        End If
    End If
#Else
        If CmbTecla.TexT = "" Or (CmbTecla.TexT = "0") Then ' And OptAlta(0).Value = True) Then
            MsgBox CargaCadena(1377) 'CargaCadena(645)  '"Introduzca Tecla"
            CmbTecla.SetFocus
            Exit Sub
        End If
#End If
    MiVendedor.seccion = CmbMostrador.TexT  'Mostradores.Mostrador(1)
        If CmbTecla.TexT = "" Then
        MiVendedor.tecla = 0
    Else
        MiVendedor.tecla = CmbTecla.TexT
    End If
    MiVendedor.Descripcion = CmbNombre.TexT
    MiVendedor.NVendedor = Text1.TexT
    
    Retorno = Alta_Vendedor(MiVendedor)
    Select Case Retorno
        Case 0
            LblInfo2.Caption = CargaCadena(527)  '"Datos Añadidos. Envíe Modificaciones"
            refresca_Nombres
            'refresca_teclas
            Refresca_Datos
        Case 1
            LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados. Envíe Modificaciones"
            refresca_Nombres
            'refresca_teclas
            Refresca_Datos
        Case 20, 3
            MsgBox CargaCadena(1378) 'CargaCadena(182)  '"Tecla ya asignada a otro vendedor"
            Text1.SetFocus
    End Select
    

End Sub

Private Sub CmdBorrar_Click()
    Dim Interruptor As Boolean
    
    Dim Registro As New RecordNet
    Dim MiVendedor As DB_Vendedor
    If Text1.TexT = "" Then Exit Sub
    MiVendedor.NVendedor = Text1.TexT
    If Baja_vendedor(MiVendedor) = 0 Then
        LblInfo2.Caption = CargaCadena(556)  '"Datos Borrados.Envíe Modificaciones"
        refresca_Nombres
        'refresca_teclas
        Refresca_Datos
    End If
End Sub

Private Sub CmdListado_Click()
     Me.Enabled = False
     Aux_Ejecuta "rpvendor", ""
     Me.Enabled = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CambiarIdioma()
    
    LblInfo.Caption = CargaCadena(524)
    Label2(3).Caption = CargaCadena(18)
    Label2(0).Caption = CargaCadena(552)
    Label2(1).Caption = CargaCadena(302)
    Label2(2).Caption = CargaCadena(78)
    CmdListado.Caption = CargaCadena(490)
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
    FrmVendedor2.Caption = CargaCadena(25)
End Sub
Private Sub Form_Load()
    Dim n As Integer
    
    Me.Icon = Form2.Icon

#If BALSAM = True Then
    CmbTecla.Clear
    For n = 1 To 9999
        CmbTecla.AddItem CStr(n)
    Next n
#Else
    CmbTecla.Clear
    For n = 1 To 30
        CmbTecla.AddItem CStr(n)
    Next n
#End If

    CambiarIdioma
    refresca_Nombres
    Refresca_Mostrador
    If Text1.ListCount > 0 Then
        Text1.ListIndex = 0
    End If
End Sub

Private Sub Text1_Click()
    Refresca_Datos
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Text1.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then Text1.Locked = True
    If Len(Text1.TexT) >= 6 Then Text1.Locked = True
End Sub
Private Function Busca_Mostrador(LaCadena As String, elDato As Integer) As Boolean
    Dim Buffer As String
    Dim bucle As Long
    Dim Encontrado As Boolean
    Encontrado = False
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            If Val(Buffer) = elDato Then Encontrado = True
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Busca_Mostrador = Encontrado
End Function
Private Function Desglosa_mostradores(LaCadena As String) As TipoMostradores
    Dim LosMostradores As TipoMostradores
    Dim Buffer As String
    Dim bucle As Long
    Dim cantidad As Integer
    cantidad = 0
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            cantidad = cantidad + 1
            ReDim Preserve LosMostradores.Mostrador(cantidad + 1)
            LosMostradores.Mostrador(cantidad) = Buffer
            LosMostradores.cantidad = cantidad
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Desglosa_mostradores = LosMostradores
End Function

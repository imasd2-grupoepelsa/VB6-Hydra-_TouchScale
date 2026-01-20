VERSION 5.00
Begin VB.Form FrmLineaComandos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de la Línea de comandos"
   ClientHeight    =   4320
   ClientLeft      =   6990
   ClientTop       =   2115
   ClientWidth     =   7335
   Icon            =   "FrmLineaComandos.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdDefecto 
      Caption         =   "Valores por defecto"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   3720
      Width           =   1692
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
      Height          =   492
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Width           =   1092
   End
   Begin VB.CommandButton CmdSalir 
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
      Height          =   492
      Left            =   6000
      TabIndex        =   3
      Top             =   3720
      Width           =   1212
   End
   Begin VB.CommandButton Cmdaceptar 
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
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Height          =   3492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7092
      Begin VB.TextBox TxtAdicional 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   2640
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2520
         Width           =   4212
      End
      Begin VB.TextBox TxtAdicional 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   2640
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2280
         Width           =   4212
      End
      Begin VB.TextBox TxtAdicional 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   4212
      End
      Begin VB.TextBox TxtAdicional 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1800
         Width           =   4212
      End
      Begin VB.CheckBox ChkSalir 
         Caption         =   "Salir del Programa Tras Ejecutar el Comando"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   5532
      End
      Begin VB.ComboBox cmbaccion 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   5412
      End
      Begin VB.ComboBox CmbTexto 
         Height          =   288
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   4212
      End
      Begin VB.Label Label1 
         Caption         =   "Indique con símbolos ## el lugar donde se indicará el número de tienda"
         Height          =   492
         Left            =   2640
         TabIndex        =   20
         Top             =   480
         Width           =   4212
      End
      Begin VB.Label lblAdicional 
         Caption         =   "Fichero 4"
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   2292
      End
      Begin VB.Label lblAdicional 
         Caption         =   "Fichero 3"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2292
      End
      Begin VB.Label lblAdicional 
         Caption         =   "Fichero 2"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   2292
      End
      Begin VB.Label lblAdicional 
         Caption         =   "Fichero 1"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2292
      End
      Begin VB.Label Label3 
         Caption         =   "Acción "
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Definición del Comando"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2172
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
         ForeColor       =   &H80000002&
         Height          =   252
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   4212
      End
      Begin VB.Label LblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "Nuevo Dato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2292
      End
   End
End
Attribute VB_Name = "FrmLineaComandos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TipoDefinicion
    texto As String
    accion As String
    Adicional As String
End Type
Private Definiciones() As TipoDefinicion
Private CuentaDefiniciones As Integer
Private QL As Boolean

Private Sub cmbaccion_LostFocus()
RefrescaDatos
End Sub

'\\\\\\\\\\\\\\\\\\\\\
'/////////////////////
Private Sub Form_Load()
Dim bucle As Integer
    If HayMulti Then
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If
    CambiarIdioma
    Me.Icon = Form2.Icon
    If Dir(Miruta & "\cmdlinetouch.cfg") = "" Then
        CrearParametrosDefecto
    End If
    For bucle = 0 To 3
        lblAdicional(bucle).Visible = False
        TxtAdicional(bucle).Visible = False
        TxtAdicional(bucle).TexT = ""
    Next bucle
    ChkSalir.Value = vbChecked
    refresca_CmbAccion
    Refresca_Definiciones
End Sub
Private Sub CambiarIdioma()
Dim bucle As Integer
    FrmLineaComandos.Caption = CargaCadena(682)
    LblInfo.Caption = CargaCadena(524)
    Label2.Caption = CargaCadena(744)
    Label3.Caption = CargaCadena(743)
    For bucle = 0 To 3
        lblAdicional(bucle).Caption = CargaCadena(745) & " " & (bucle + 1)
    Next bucle
    ChkSalir.Caption = CargaCadena(746)
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdDefecto.Caption = CargaCadena(747)
    CmdSalir.Caption = CargaCadena(7)
End Sub
Private Sub cmbaccion_Click()
    Dim bucle As Integer
    If Val(left(cmbaccion.TexT, 3)) = 17 Or Val(left(cmbaccion.TexT, 3)) = 19 _
    Or Val(left(cmbaccion.TexT, 3)) = 34 Or Val(left(cmbaccion.TexT, 3)) = 38 Then
    'If Val(left(cmbaccion.TexT, 3)) = 15 Or Val(left(cmbaccion.TexT, 3)) = 7 Then
        For bucle = 0 To 3
            lblAdicional(bucle).Visible = True
            TxtAdicional(bucle).Visible = True
        Next bucle
    Else
        For bucle = 0 To 3
            lblAdicional(bucle).Visible = False
            TxtAdicional(bucle).Visible = False
        Next bucle
    End If
    If Val(left(cmbaccion.TexT, 3)) = 36 Then
        lblAdicional(0).Visible = True
        TxtAdicional(0).Visible = True
    End If
    RefrescaDatos
End Sub

Private Sub CmbTexto_Click()
RefrescaDatos
End Sub

Private Sub CmbTexto_LostFocus()
    RefrescaDatos
End Sub
Private Sub graba_Fichero()
    Dim Fichero As Integer
    Dim bucle As Integer
    Fichero = FreeFile
    Open Miruta & "\cmdlinetouch.cfg" For Output As #Fichero
    For bucle = 1 To CuentaDefiniciones
        Print #Fichero, Definiciones(bucle).texto
        Print #Fichero, Definiciones(bucle).accion
        Print #Fichero, Definiciones(bucle).Adicional
    Next bucle
    Close #Fichero
End Sub
Private Sub Cmdaceptar_Click()
    
    Dim bucle As Integer
    Dim B2 As Integer
    Dim Buffer As String
    Dim eNCONTRADO As Boolean
    eNCONTRADO = False
    If Val(left(cmbaccion.TexT, 3)) = 37 Then
        If InStr(1, CmbTexto.TexT, "?") = 0 Then
            MsgBox CargaCadena(1036), vbCritical
            Exit Sub
        End If
    End If
    If Trim(CmbTexto.TexT) = "" Then Exit Sub
    If Trim(cmbaccion.TexT) = "" Then
        MsgBox CargaCadena(748)
        Exit Sub
    End If
    If Val(left(cmbaccion.TexT, 3)) = 36 Or Val(left(cmbaccion.TexT, 3)) = 17 Or _
    Val(left(cmbaccion.TexT, 3)) = 19 Or Val(left(cmbaccion.TexT, 3)) = 34 Or Val(left(cmbaccion.TexT, 3)) = 38 Then
    'If Val(left(cmbaccion.TexT, 3)) = 7 Or Val(left(cmbaccion.TexT, 3)) = 15 Then
        If Trim(TxtAdicional(0).TexT) = "" And Trim(TxtAdicional(1).TexT) = "" And Trim(TxtAdicional(2).TexT) = "" And Trim(TxtAdicional(3).TexT) = "" Then
            MsgBox CargaCadena(749)  '"Escriba los nombres de los ficheros"
            Exit Sub
        End If
    End If
    If ChkSalir.Value = vbChecked Then
            Buffer = "Y"
        Else
            Buffer = "N"
        End If
        If Val(left(cmbaccion.TexT, 3)) = 17 Or Val(left(cmbaccion.TexT, 3)) = 36 _
        Or Val(left(cmbaccion.TexT, 3)) = 19 Or Val(left(cmbaccion.TexT, 3)) = 34 Or Val(left(cmbaccion.TexT, 3)) = 38 Then
        'If Val(left(cmbaccion.TexT, 3)) = 7 Or Val(left(cmbaccion.TexT, 3)) = 15 Then
            For B2 = 0 To 3
                If Trim(TxtAdicional(B2).TexT) <> "" Then
                    If B2 > 0 Then Buffer = Buffer & "\"
                    Buffer = Buffer & TxtAdicional(B2).TexT
                End If
            Next B2
    End If
    For bucle = 1 To CuentaDefiniciones
        If CmbTexto.TexT = Definiciones(bucle).texto Then
            eNCONTRADO = True
            Exit For
        End If
    Next bucle
    If eNCONTRADO Then
        LblInfo2.Caption = CargaCadena(751)  '"Datos Modificados"
        If Buffer = "Y" And InStr(1, Definiciones(bucle).texto, "(exit)") = 0 Then
            CmbTexto.TexT = CmbTexto.TexT & " (exit)"
        End If
        Definiciones(bucle).texto = CmbTexto.TexT
        Definiciones(bucle).accion = cmbaccion.TexT
        Definiciones(bucle).Adicional = Buffer
    Else
        LblInfo2.Caption = CargaCadena(752)  '"Datos Añadidos"
        CuentaDefiniciones = CuentaDefiniciones + 1
        ReDim Preserve Definiciones(CuentaDefiniciones + 1)
        If Buffer = "Y" And InStr(1, Definiciones(CuentaDefiniciones).texto, "(exit)") = 0 Then
            CmbTexto.TexT = CmbTexto.TexT & " (exit)"
        End If
        Definiciones(CuentaDefiniciones).texto = CmbTexto.TexT
        Definiciones(CuentaDefiniciones).accion = cmbaccion.TexT
        Definiciones(CuentaDefiniciones).Adicional = Buffer
    End If
    graba_Fichero
    
    CmbTexto.TexT = ""
    For bucle = 0 To 3
        TxtAdicional(bucle).TexT = ""
        TxtAdicional(bucle).Visible = False
        lblAdicional(bucle).Visible = False
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        refresca_CmbAccion
        Refresca_Definiciones
    Next bucle
    
End Sub

Private Sub CmdBorrar_Click()
    Dim Def2() As TipoDefinicion
    Dim bucle As Integer
    Dim Contador As Integer
    Dim eNCONTRADO As Boolean
    eNCONTRADO = False
    If Trim(CmbTexto.TexT) = "" Then Exit Sub
    For bucle = 1 To CuentaDefiniciones
        If CmbTexto.TexT = Definiciones(bucle).texto Then
            eNCONTRADO = True
            Exit For
        End If
    Next bucle
    
    If eNCONTRADO Then
        Contador = 1
        ReDim Def2(CuentaDefiniciones)
        For bucle = 1 To CuentaDefiniciones
            If Definiciones(bucle).texto <> CmbTexto.TexT Then
                Def2(Contador).accion = Definiciones(bucle).accion
                Def2(Contador).texto = Definiciones(bucle).texto
                Def2(Contador).Adicional = Definiciones(bucle).Adicional
                Contador = Contador + 1
            End If
        Next bucle
        CuentaDefiniciones = CuentaDefiniciones - 1
        ReDim Definiciones(CuentaDefiniciones)
        For bucle = 1 To CuentaDefiniciones
            Definiciones(bucle).accion = Def2(bucle).accion
            Definiciones(bucle).texto = Def2(bucle).texto
            Definiciones(bucle).Adicional = Def2(bucle).Adicional
        Next bucle
        LblInfo2.Caption = CargaCadena(528)  '"Datos Borrados"
        graba_Fichero
        refresca_CmbAccion
        Refresca_Definiciones
        For bucle = 0 To 3
            TxtAdicional(bucle).TexT = ""
            TxtAdicional(bucle).Visible = False
            lblAdicional(bucle).Visible = False
        Next bucle
    End If
End Sub

Private Sub CmdDefecto_Click()
    Dim Res
    Res = MsgBox(CargaCadena(750), vbYesNo + vbExclamation)
    If Res = vbYes Then
        CrearParametrosDefecto
        Unload Me
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub RefrescaDatos()
    Dim bucle As Integer
    Dim B2 As Integer
    Dim b3 As Integer
    Dim Ctdor As Integer
    Dim eNCONTRADO As Boolean
    If QL = True Then Exit Sub
    QL = True
    eNCONTRADO = False
    If CmbTexto.TexT = "" Then
        LblInfo.BackColor = vbWhite
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        refresca_CmbAccion
        For B2 = 0 To 3
            lblAdicional(B2).Visible = False
            TxtAdicional(B2).Visible = False
            TxtAdicional(B2).TexT = ""
        Next B2
        ChkSalir.Value = vbChecked
    Else
        If left(CmbTexto.TexT, 1) <> "/" Then
            CmbTexto.TexT = "/" & CmbTexto.TexT
        End If
        CmbTexto.TexT = StrConv(CmbTexto.TexT, vbUpperCase)
        For bucle = 1 To CuentaDefiniciones
            If Definiciones(bucle).texto = CmbTexto.TexT Then
                LblInfo.Caption = CargaCadena(525)  '"Modificación"
                LblInfo.BackColor = vbYellow
                CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
                CmbTexto.TexT = Definiciones(bucle).texto
                cmbaccion.ListIndex = Val(left(cmbaccion.List(Val(left(Definiciones(bucle).accion, 3)) - 1), 3)) - 1
                If left(Definiciones(bucle).Adicional, 1) = "Y" Then
                    ChkSalir.Value = vbChecked
                Else
                    ChkSalir.Value = vbUnchecked
                End If
                If Val(left(Definiciones(bucle).accion, 3)) = "17" _
                Or Val(left(Definiciones(bucle).accion, 3)) = "19" _
                Or Val(left(Definiciones(bucle).accion, 3)) = "34" _
                Or Val(left(Definiciones(bucle).accion, 3)) = "38" Then
                    Ctdor = 0
                    For B2 = 0 To 3
                        lblAdicional(B2).Visible = True
                        TxtAdicional(B2).Visible = True
                        TxtAdicional(B2).TexT = ""
                    Next B2
                    For b3 = 2 To Len(Definiciones(bucle).Adicional)
                        If Mid(Definiciones(bucle).Adicional, b3, 1) = "\" Then
                            Ctdor = Ctdor + 1
                        Else
                            TxtAdicional(Ctdor).TexT = TxtAdicional(Ctdor).TexT & Mid(Definiciones(bucle).Adicional, b3, 1)
                        End If
                        If Ctdor > 3 Then Exit For
                    Next b3
                Else
                    For B2 = 0 To 3
                        lblAdicional(B2).Visible = False
                        TxtAdicional(B2).Visible = False
                        TxtAdicional(B2).TexT = ""
                    Next B2
                    If Val(left(Definiciones(bucle).accion, 3)) = "36" Then
                        lblAdicional(0).Visible = True
                        TxtAdicional(0).Visible = True
                        TxtAdicional(0).TexT = Mid(Definiciones(bucle).Adicional, 2)
                    End If
                End If
                eNCONTRADO = True
            End If
            If eNCONTRADO = True Then Exit For
        Next bucle
        If Not eNCONTRADO Then
            LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
            LblInfo.BackColor = vbWhite
            refresca_CmbAccion
            ChkSalir.Value = vbChecked
            For B2 = 0 To 3
                lblAdicional(B2).Visible = False
                TxtAdicional(B2).Visible = False
                TxtAdicional(B2).TexT = ""
            Next B2
        End If
    End If
    QL = False
End Sub
Private Sub Refresca_Definiciones()
    Dim bucle As Integer
    Dim Buffer As String
    Dim Fichero As Integer
    CmbTexto.Clear
    CuentaDefiniciones = 0
    If Dir(Miruta & "\cmdlinetouch.cfg") <> "" Then
        Fichero = FreeFile()
        Open Miruta & "\cmdlinetouch.cfg" For Input As #Fichero
        If Not EOF(Fichero) Then
            Do Until EOF(Fichero)
                CuentaDefiniciones = CuentaDefiniciones + 1
                ReDim Preserve Definiciones(CuentaDefiniciones + 1)
                Line Input #Fichero, Definiciones(CuentaDefiniciones).texto
                Line Input #Fichero, Definiciones(CuentaDefiniciones).accion
                Line Input #Fichero, Definiciones(CuentaDefiniciones).Adicional
            Loop
        End If
        '''
        'cmbaccion.Clear
        '''
        'For bucle = 1 To CuentaDefiniciones
        '    CmbTexto.AddItem Definiciones(bucle).texto
        '    If Definiciones(bucle).Adicional = "Y" Then
        '        cmbaccion.AddItem Definiciones(bucle).accion & " (exit)"
        '    Else
        '        cmbaccion.AddItem Definiciones(bucle).accion
        '    End If
        'Next bucle
        Close #Fichero
    End If
End Sub
Private Sub refresca_CmbAccion()
    cmbaccion.Clear
    cmbaccion.AddItem "1   " & CargaCadena(764)
    cmbaccion.AddItem "2   " & CargaCadena(151)
    cmbaccion.AddItem "3   " & CargaCadena(225) & " " & CargaCadena(9)
    cmbaccion.AddItem "4   " & CargaCadena(225) & " " & CargaCadena(677)
    cmbaccion.AddItem "5   " & CargaCadena(225) & " " & CargaCadena(678)
    cmbaccion.AddItem "6   " & CargaCadena(225) & " " & CargaCadena(11)
    cmbaccion.AddItem "7   " & CargaCadena(225) & " " & CargaCadena(212)
    cmbaccion.AddItem "8   " & CargaCadena(225) & " " & CargaCadena(619)
    cmbaccion.AddItem "9   " & CargaCadena(225) & " " & CargaCadena(12)
    cmbaccion.AddItem "10  " & CargaCadena(225) & " " & CargaCadena(14)
    cmbaccion.AddItem "11  " & CargaCadena(225) & " " & CargaCadena(21)
    cmbaccion.AddItem "12  " & CargaCadena(225) & " " & CargaCadena(23)
    cmbaccion.AddItem "13  " & CargaCadena(225) & " " & CargaCadena(13)
    cmbaccion.AddItem "14  " & CargaCadena(225) & " " & CargaCadena(22)
    cmbaccion.AddItem "15  " & CargaCadena(161)
    cmbaccion.AddItem "16  " & CargaCadena(1264) 'Recepción de Fichas de Vacuno
    cmbaccion.AddItem "17  " & CargaCadena(685)
    cmbaccion.AddItem "18  " & CargaCadena(1265) 'Importar Fichas de Vacuno
    cmbaccion.AddItem "19  " & CargaCadena(686)
    cmbaccion.AddItem "20  " & CargaCadena(1197) & " " & CargaCadena(580)
    cmbaccion.AddItem "21  " & CargaCadena(1197) & " " & CargaCadena(12)
    cmbaccion.AddItem "22  " & CargaCadena(1197) & " " & CargaCadena(13)
    cmbaccion.AddItem "23  " & CargaCadena(1197) & " " & CargaCadena(14)
    cmbaccion.AddItem "24  " & CargaCadena(1197) & " " & CargaCadena(21)
    cmbaccion.AddItem "25  " & CargaCadena(1197) & " " & CargaCadena(23)
    cmbaccion.AddItem "26  " & CargaCadena(1197) & " " & CargaCadena(22)
    cmbaccion.AddItem "27  " & CargaCadena(1197) & " " & CargaCadena(406)
    cmbaccion.AddItem "28  " & CargaCadena(1197) & " " & CargaCadena(410)
    cmbaccion.AddItem "29  " & CargaCadena(1197) & " " & CargaCadena(212)
    cmbaccion.AddItem "30  " & CargaCadena(1197) & " " & CargaCadena(619)
    cmbaccion.AddItem "31  " & CargaCadena(1197) & " " & CargaCadena(24)
    cmbaccion.AddItem "32  " & CargaCadena(1197) & " " & CargaCadena(8)
    cmbaccion.AddItem "33  " & CargaCadena(225) & " " & CargaCadena(8)
    cmbaccion.AddItem "34  " & CargaCadena(236) & " + " & CargaCadena(151)
    cmbaccion.AddItem "35  " & CargaCadena(101) 'Programación de Fecha y Hora
    cmbaccion.AddItem "36  " & CargaCadena(1266) 'Importa datos de balanza y genera fichero
    cmbaccion.AddItem "37  " & CargaCadena(927) 'Borra una sección
    cmbaccion.AddItem "38  " & CargaCadena(1267) 'Exportar Totales Históricos
    cmbaccion.AddItem "39  " & CargaCadena(1268) & " + HydraTouch" 'Borra datos de Balanzas + HydraTouch
    cmbaccion.AddItem "40  " & CargaCadena(1269) 'Backup Bases de Datos MultiTienda
    cmbaccion.AddItem "41  " & CargaCadena(1095)
    cmbaccion.AddItem "42  " & CargaCadena(1093)
    cmbaccion.AddItem "43  " & CargaCadena(1094)
    cmbaccion.AddItem "44  " & CargaCadena(1175) 'Importar GLOBAL.DAT
    cmbaccion.AddItem "45  " & CargaCadena(1175) & " + " & CargaCadena(151) 'Importar GLOBAL.DAT + envío modificaciones"
    cmbaccion.AddItem "46  " & CargaCadena(1197) & " " & CargaCadena(1114)
    cmbaccion.AddItem "47  " & CargaCadena(1197) & " " & CargaCadena(1115)
    cmbaccion.AddItem "48  " & CargaCadena(225) & " " & CargaCadena(1114)
    cmbaccion.AddItem "49  " & CargaCadena(225) & " " & CargaCadena(1115)
    cmbaccion.AddItem "50  " & CargaCadena(514) & " + " & CargaCadena(151) '"Recoger totales + envío modificaciones"
    cmbaccion.AddItem "51  " & CargaCadena(1270) ' "Recoger totales sin borrado
    cmbaccion.AddItem "52  " & "Erase totals Data Tables."
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub TxtAdicional_KeyPress(Index As Integer, KeyAscii As Integer)
    TxtAdicional(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub
    If Chr(KeyAscii) = ":" Or Chr(KeyAscii) = "\" Then TxtAdicional(Index).Locked = True
End Sub

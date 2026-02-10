VERSION 5.00
Begin VB.Form FrmMostrador2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "---"
   ClientHeight    =   5385
   ClientLeft      =   4980
   ClientTop       =   2595
   ClientWidth     =   4710
   Icon            =   "Frmmostrador2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4710
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   22
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CheckBox ChkAhoSS 
      Caption         =   "Self-Service."
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TextPeco 
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox TextPathSec 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.ComboBox CmbCOM 
      Height          =   315
      ItemData        =   "Frmmostrador2.frx":000C
      Left            =   3120
      List            =   "Frmmostrador2.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3480
      TabIndex        =   9
      Top             =   4920
      Width           =   1092
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
      TabIndex        =   7
      Top             =   4920
      Width           =   1092
   End
   Begin VB.TextBox TxtMaquina 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.ComboBox CmbMostrador 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "Frmmostrador2.frx":0078
      Left            =   3240
      List            =   "Frmmostrador2.frx":007A
      TabIndex        =   1
      Top             =   840
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox CmbEnviar 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "Frmmostrador2.frx":007C
      Left            =   2400
      List            =   "Frmmostrador2.frx":008F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2172
   End
   Begin VB.Label Label4 
      Caption         =   "Descripcion:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label LabelPeco 
      Caption         =   "TCP Port:"
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
      Left            =   1560
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "IP de este sistema"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Puerto de Comunicaciones"
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
      TabIndex        =   15
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   4455
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
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4452
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2652
   End
   Begin VB.Label Label2 
      Caption         =   "0"
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2532
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2172
   End
End
Attribute VB_Name = "FrmMostrador2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QL1 As Boolean




Private Sub CmbCOM_Click()    '2.0.a para controlar eco a través de IP
    If (CmbEnviar.ListIndex = 3) Then
        If CmbCOM.ListIndex <> 0 Then
            tXTip.TexT = ""
            LabelPeco.Visible = True
            tXTip.Enabled = False
            TextPeco.TexT = ""
            TextPeco.Visible = True
            TextPeco.Enabled = False
        Else
            LabelPeco.Visible = True
            TextPeco.Visible = True
            tXTip.Enabled = True
            TextPeco.Enabled = True
        End If
    End If

End Sub

Private Sub CmbEnviar_Click()
    Dim MyOK As Boolean
    MyOK = True

    Label2(3).Visible = False
    CmbCOM.Visible = False
    If CmbEnviar.ListIndex = 2 Or CmbEnviar.ListIndex = 3 Then
        If Not UsaModem Then
            Label2(3).Visible = True
            CmbCOM.Visible = True
        End If
    End If
    If QL1 = True Then Exit Sub
    '*******************************
    ' corrige el valor del combo
    ' si el usuario ha elegido
    ' un tipo de sección incorrecto
    '*******************************
    If CmbEnviar.ListIndex = 0 And Not UsaGamaAlta Then MyOK = False
    If CmbEnviar.ListIndex = 1 And Not usaSC10 Then MyOK = False
    If CmbEnviar.ListIndex = 2 And Not UsaECO1 Then MyOK = False
    If CmbEnviar.ListIndex = 3 And Not UsaECO2 Then MyOK = False
    If CmbEnviar.ListIndex = 4 And Not UsaEcoPrint Then MyOK = False
    If CmbEnviar.ListIndex = 5 And Not (usaSC10 And UsaGamaAlta) Then MyOK = False
    '********************************
    ' Sitúa un valor por defecto
    '********************************
    If MyOK = False Then
        QL1 = True
        If UsaEcoPrint Then CmbEnviar.ListIndex = 4
        If UsaECO2 Then CmbEnviar.ListIndex = 3
        If UsaECO1 Then CmbEnviar.ListIndex = 2
        If usaSC10 Then CmbEnviar.ListIndex = 1
        If UsaGamaAlta Then CmbEnviar.ListIndex = 0
        QL1 = False
    End If
    If ((CmbEnviar.ListIndex = 0 Or CmbEnviar.ListIndex = 5) And MultiEuroscale) Or ((CmbEnviar.ListIndex = 3) And (CmbCOM.ListIndex = 0)) Then  '2.0.a se añade usaeco2
        Label3.Visible = True
        tXTip.Visible = True
        If CmbEnviar.ListIndex = 3 Then
            LabelPeco.Visible = True
            TextPeco.Visible = True
        Else
            LabelPeco.Visible = False
            TextPeco.Visible = False

        End If
    Else
        Label3.Visible = False
        tXTip.Visible = False
        LabelPeco.Visible = False
        TextPeco.Visible = False

    End If
End Sub

Private Sub CmbMostrador_click()
'******************
' busca los datos del mostrador
    Refresca_Datos
End Sub

Private Sub CmbMostrador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    CmbMostrador.Locked = False
    If KeyAscii < 48 Or KeyAscii > 57 Then CmbMostrador.Locked = True
    If Len(CmbMostrador.TexT) >= 3 Then
        If Len(CmbMostrador.TexT) > 3 Then CmbMostrador.TexT = left(CmbMostrador.TexT, 3)
        CmbMostrador.Locked = True
    End If
End Sub
Private Sub CmbMostrador_LostFocus()
    Refresca_Datos
End Sub
Private Sub Cmdaceptar_Click()
    Dim salir As Boolean
    Dim MiMostrador As DB_mostrador
    Dim nf As Integer
    Dim ncont As Integer

    '***************************************
    ' Introducción o modificación de datos
    '****************************************
    '/////INI VERIFICACIONES
    MiMostrador.DireccionIP = ""

    If Dir(App.Path & "\plrm.txt") <> "" Then
        nf = FreeFile()
        If nContSecPlrm = 0 Then
            nContSecPlrm = nContSecPlrm + 1
            nSecPLRM(nContSecPlrm) = Val(CmbMostrador.TexT)
            sPathSecPlrm(nContSecPlrm) = Trim(TextPathSec.TexT)
            If Right(sPathSecPlrm(nContSecPlrm), 1) <> "\" Then sPathSecPlrm(nContSecPlrm) = sPathSecPlrm(nContSecPlrm) & "\"
            Open App.Path & "\plrm.txt" For Output As #nf
            Print #nf, Format(Val(CmbMostrador.TexT), "00") & Trim(TextPathSec.TexT)
            Close #nf
        Else
            For ncont = 1 To nContSecPlrm
                If Val(nSecPLRM(ncont)) = Val(CmbMostrador.TexT) Then
                    sPathSecPlrm(ncont) = Trim(TextPathSec.TexT)
                    Exit For
                End If
            Next ncont
            If ncont > nContSecPlrm Then
                nContSecPlrm = nContSecPlrm + 1
                nSecPLRM(nContSecPlrm) = Val(CmbMostrador.TexT)
                sPathSecPlrm(nContSecPlrm) = Trim(TextPathSec.TexT)
                If Right(sPathSecPlrm(nContSecPlrm), 1) <> "\" Then sPathSecPlrm(nContSecPlrm) = sPathSecPlrm(nContSecPlrm) & "\"
                nf = FreeFile()
                Open App.Path & "\plrm.txt" For Output As #nf
                For ncont = 1 To nContSecPlrm
                    Print #nf, Format(nSecPLRM(ncont), "00") & sPathSecPlrm(ncont)
                Next ncont
                Close #nf
            End If
        End If

    End If

    If MultiEuroscale Or (CmbEnviar.ListIndex = 3) Then    '2.0.a ip para eco...
        If CmbEnviar.ListIndex = 0 Or CmbEnviar.ListIndex = 5 Or CmbEnviar.ListIndex = 3 Then    '2.0.a Ip eco...
            If tXTip.TexT = "" And CmbEnviar.ListIndex <> 3 Then    '2.0.a ip eco
                MsgBox CargaCadena(577), vbCritical  '"Falta la dirección IP"
                tXTip.SetFocus
                Exit Sub
            Else
                MiMostrador.DireccionIP = tXTip.TexT
            End If
        End If
    End If
    If CmbMostrador.TexT = "" Then
        MsgBox CargaCadena(559)  '"Introduzca un número de Mostrador"
        Exit Sub
    End If
    If TxtMaquina.TexT = "" Then
        MsgBox CargaCadena(560)  '"Introduzca un número de Sección"
        TxtMaquina.SetFocus
        Exit Sub
    End If
    '***********************************
    ' para Euroscale, número de sección máquina
    ' máximo es 60
    '***********************************
    If (TxtMaquina.TexT > 60) And _
       ((CmbEnviar.TexT = CmbEnviar.List(0) Or CmbEnviar.TexT = CmbEnviar.List(5))) Then
        MsgBox CargaCadena(561)  '"Euroscale sólo admite números de sección entre 1 y 60"
        TxtMaquina.SetFocus
        Exit Sub
    End If
    '/////FIN VERIFICACIONES
    MiMostrador.PuertoCOM = CmbCOM.ListIndex
    MiMostrador.NMostrador = CmbMostrador.TexT
    MiMostrador.NSeccion = TxtMaquina.TexT
    MiMostrador.name = Text1.TexT
    MiMostrador.Descripcion = Text2.TexT
    Select Case CmbEnviar.ListIndex
    Case 0
        MiMostrador.Tipo = 0
    Case 1
        MiMostrador.Tipo = 1
    Case 2
        MiMostrador.Tipo = 3    'usaeco2
    Case 3
        MiMostrador.Tipo = 4
    Case 4
        MiMostrador.Tipo = 5
    Case 5
        MiMostrador.Tipo = 2
    End Select

    salir = False
    If CmbMostrador.Enabled = True Then CmbMostrador.SetFocus
    Select Case Alta_Mostrador(MiMostrador)
    Case 7
        MsgBox CargaCadena(562)  '"El número de mostrador no puede ser superior a 3 dígitos"
        Exit Sub
    Case 8
        MsgBox CargaCadena(563)  '"El número de sección no puede ser superior a 2 dígitos"
        TxtMaquina.SetFocus
        Exit Sub
    Case 9
        MsgBox CargaCadena(564)  '"La descripción no puede tener más de 21 caracteres"
        Exit Sub
    Case 50
        MsgBox CargaCadena(1116), vbCritical
        Exit Sub
    Case 51
        MsgBox CargaCadena(1117), vbCritical
        Exit Sub

    Case 0
        If CmbMostrador.Enabled = False Then salir = True

        LblInfo2.Caption = CargaCadena(527)  '"Datos Añadidos.Envíe Modificaciones"
        FrmSeccionyMaquina.Arbol.Nodes.Add "sistema", tvwChild, "S" & MiMostrador.NMostrador, MiMostrador.NMostrador, "seccion"
    Case 1
        If CmbMostrador.Enabled = False Then salir = True
        LblInfo2.Caption = CargaCadena(528)  '"Datos Modificados.Envíe modificaciones"

    End Select
    If Dir(App.Path & "\plrm.txt") <> "" Then
        nf = FreeFile()
        Open App.Path & "\plrm.txt" For Output As #nf
        For ncont = 1 To nContSecPlrm
            Print #nf, Format(nSecPLRM(ncont), "00") & sPathSecPlrm(ncont)
        Next ncont
        Close #nf
    End If
    If salir Then Unload Me
End Sub

Private Sub Command1_Click()
    Dim MiMostrador As DB_mostrador
    '*****************************
    ' Orden de borrado de datos
    ' asociados
    '*****************************
    If CmbMostrador.TexT = "" Then Exit Sub
    MiMostrador.NMostrador = CmbMostrador.TexT
    Select Case Baja_Mostrador(MiMostrador)
    Case 0
        LblInfo2.Caption = CargaCadena(556)  '"Datos Eliminados.Envíe Modificaciones"
        Refresca_Mostrador
        FrmSeccionyMaquina.Refresca_Arbol
        CmbMostrador.TexT = ""
        Refresca_Datos
    Case 10
        MsgBox CargaCadena(557)  '"No se puede borrar el Mostrador.Tiene artículos asociados"
    Case 11
        MsgBox CargaCadena(558)  '"No se puede borrar el mostrador.Tiene Equipos Asociados"
    End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim nf As Integer

    '*****************
    ' Aspecto
    '*****************
    CmbCOM.ListIndex = 0
    '2.0.a
    LabelPeco.Visible = False
    TextPeco.Visible = False
    '''
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    CambiarIdioma
    '***************************
    ' Sitúa en el combo los
    ' tipos de secciones disponibles
    '***************************
    If Dir(App.Path & "\plrm.txt") <> "" Then
        nContSecPlrm = 0
        nf = FreeFile()
        Open App.Path & "\plrm.txt" For Input As #nf
        Do While Not EOF(nf)
            nContSecPlrm = nContSecPlrm + 1
            Line Input #nf, sPathSecPlrm(nContSecPlrm)
            sPathSecPlrm(nContSecPlrm) = Trim(sPathSecPlrm(nContSecPlrm))
            If Right(sPathSecPlrm(nContSecPlrm), 1) <> "\" Then sPathSecPlrm(nContSecPlrm) = sPathSecPlrm(nContSecPlrm) & "\"
            nSecPLRM(nContSecPlrm) = Val(Mid(sPathSecPlrm(nContSecPlrm), 1, 2))
            sPathSecPlrm(nContSecPlrm) = Mid(sPathSecPlrm(nContSecPlrm), 3)
        Loop
        Close #nf
        TextPathSec.Visible = True
    End If

    If usaSC10 And TipoC9 Then
        CmbEnviar.List(1) = "C9/C10"
    Else
        CmbEnviar.List(1) = "SC10"
    End If
    'If UsaECO1 And TipoV8 Then
    '    CmbEnviar.List(2) = "V-8"
    'Else
    '    CmbEnviar.List(2) = "Red Eco 1"
    'End If
    If UsaEcoPrint Then
        Select Case TipoEcoPrint
        Case 0
            CmbEnviar.List(4) = "IV4 STD"
        Case 1
            CmbEnviar.List(4) = "IV4 CE"
        Case 2
            CmbEnviar.List(4) = "IV4 TQ"
        Case 3
            CmbEnviar.List(4) = "IV4 LCD"
        End Select
    End If
    '********************************
    ' sitúa un valor por defecto
    '********************************
    If UsaEcoPrint Then CmbEnviar.TexT = CmbEnviar.List(4)
    If UsaECO2 Then CmbEnviar.TexT = CmbEnviar.List(3)
    If UsaECO1 Then CmbEnviar.TexT = CmbEnviar.List(2)
    If usaSC10 Then CmbEnviar.TexT = CmbEnviar.List(1)
    If UsaGamaAlta Then CmbEnviar.TexT = CmbEnviar.List(0)
    CmbEnviar.Enabled = False
    If NumeroSistemas > 1 Then
        CmbEnviar.Enabled = True
    End If
    '2.0.36
    'If AhorraMas And usaSC10 Then
    '    ChkAhoSS.Visible = True
    'Else
    '    ChkAhoSS.Visible = False
    'End If
    ChkAhoSS.Visible = False
    '''''''''''''''''''''''''''''

    '''''''
    Refresca_Mostrador
    Refresca_Datos
End Sub
Private Sub Refresca_Datos()

    Dim Registro As New RecordNet
    Dim nf, ncont As Integer


    '2.0.a
    If CmbEnviar.ListIndex = 3 Then
        LabelPeco.Visible = True
        TextPeco.Visible = True
    Else
        LabelPeco.Visible = False
        TextPeco.Visible = False

    End If
    '*********************************
    ' prepara el formulario para
    ' aceptar un dato nuevo
    '*********************************
    TxtMaquina.TexT = ""
    Text1.TexT = ""
    Text2.TexT = ""
    tXTip.TexT = ""
    LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
    LblInfo.BackColor = vbWhite
    CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
    '1.0.0.4
    'TxtMaquina.Enabled = True
    ''''''''
    If NumeroSistemas > 1 Then
        CmbEnviar.Enabled = True
    End If
    If IsNumeric(CmbMostrador.TexT) Then

        Registro.OpenRecordset ("select * from SECCION where borrado=false and codi_ident=" & CmbMostrador.TexT)
        With Registro
            If Not .EOF Then
                '2.0.36
                'If AhorraMas And usaSC10 And .Fields("com") = 255 Then
                '    ChkAhoSS.Value = vbChecked
                'Else
                '    ChkAhoSS.Value = vbUnchecked
                'End If
                '''''''''''''''''''''''''''''''''
                '********************************
                ' si el mostrador ya existe,
                ' pone los datos del mostrador
                '********************************
                CmbMostrador.TexT = Format(.Fields("codi_ident"), "000")
                TxtMaquina.TexT = Format(.Fields("secc_Maqui"), "00")
                Text1.TexT = .Fields("nombre")
                Text2.TexT = .Fields("descripcio")
                QL1 = True
                If .Fields("enviardatos") = "GA" Then CmbEnviar.TexT = CmbEnviar.List(0)
                If .Fields("enviardatos") = "SC10" Then CmbEnviar.TexT = CmbEnviar.List(1)
                If .Fields("enviardatos") = "ECON" Then CmbEnviar.TexT = CmbEnviar.List(2)
                If .Fields("enviardatos") = "ECOL" Then CmbEnviar.TexT = CmbEnviar.List(3)
                If .Fields("enviardatos") = "ECOP" Then CmbEnviar.TexT = CmbEnviar.List(3)
                If .Fields("enviardatos") = "TODOS" Then CmbEnviar.TexT = CmbEnviar.List(5)
                CmbCOM.Visible = False
                Label2(3).Visible = False
                If .Fields("enviardatos") = "ECON" Or .Fields("enviardatos") = "ECOL" Then
                    If Not UsaModem Then
                        CmbCOM.Visible = True
                        Label2(3).Visible = True
                    End If
                End If
                If IsNumeric(.Fields("COM")) Then
                    On Error Resume Next
                    CmbCOM.ListIndex = .Fields("COM")
                    On Error GoTo 0
                Else
                    CmbCOM.ListIndex = 0
                End If
                If Not IsNull(.Fields("SEC_IP")) Then
                    tXTip.TexT = .Fields("SEC_IP")
                End If
                If CmbEnviar.ListIndex = 3 Then

                    TextPeco.TexT = CStr(.Fields("multiple"))
                End If
                QL1 = False
                LblInfo.Caption = CargaCadena(525)  '"Modificación"
                LblInfo.BackColor = vbYellow
                CmdAceptar.Caption = CargaCadena(36)  '"Modificar"
                TxtMaquina.Enabled = False
            Else
                If IsNumeric(CmbMostrador.TexT) And (Not IsNumeric(TxtMaquina.TexT)) Then
                    TxtMaquina.TexT = CmbMostrador.TexT
                End If
            End If
        End With

    End If

    If Dir(App.Path & "\plrm.txt") <> "" Then
        nContSecPlrm = 0
        nf = FreeFile()
        Open App.Path & "\plrm.txt" For Input As #nf
        Do While Not EOF(nf)
            nContSecPlrm = nContSecPlrm + 1
            Line Input #nf, sPathSecPlrm(nContSecPlrm)
            sPathSecPlrm(nContSecPlrm) = Trim(sPathSecPlrm(nContSecPlrm))
            If Right(sPathSecPlrm(nContSecPlrm), 1) <> "\" Then sPathSecPlrm(nContSecPlrm) = sPathSecPlrm(nContSecPlrm) & "\"
            nSecPLRM(nContSecPlrm) = Val(Mid(sPathSecPlrm(nContSecPlrm), 1, 2))
            sPathSecPlrm(nContSecPlrm) = Mid(sPathSecPlrm(nContSecPlrm), 3)
        Loop
        Close #nf
        For ncont = 1 To nContSecPlrm
            If Val(CmbMostrador.TexT) = Val(nSecPLRM(ncont)) Then
                TextPathSec.TexT = sPathSecPlrm(ncont)
            End If
        Next ncont
    End If
End Sub
Private Sub Refresca_Mostrador()

    Dim Registro As New RecordNet
    '***********************************
    ' Sitúa en el combo los mostradores
    ' existentes
    '***********************************

    Registro.OpenRecordset ("select codi_ident,borrado from seccion where borrado=false order by codi_ident")
    CmbMostrador.Clear
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                CmbMostrador.AddItem Format(.Fields("codi_ident"), "000")
                .Movenext
            Loop
        End If
    End With

End Sub


Private Sub Text1_Change()
    If Len(Text1.TexT) > 50 Then Text1.TexT = left(Text1.TexT, 50)
End Sub
Private Sub Text2_Change()
    If Len(Text2.TexT) > 50 Then Text2.TexT = left(Text2.TexT, 50)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Text1.Locked = False
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If Len(Text1.TexT) >= 50 Then
        Text1.Locked = True
        If Len(Text1.TexT) > 50 Then Text1.TexT = left(Text1.TexT, 50)
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Text2.Locked = False
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If Len(Text2.TexT) >= 50 Then
        Text2.Locked = True
        If Len(Text2.TexT) > 50 Then Text2.TexT = left(Text2.TexT, 50)
    End If
End Sub


'Private Sub Text2_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 48 To 59
'        Case vbKeyBack
'        Case vbKeyDelete
'        Case 46
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub





Private Sub TextPeco_KeyPress(KeyAscii As Integer)    '2.0.a
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
Private Sub textPeco_LostFocus()
    If Val(TextPeco.TexT) > 65535 Then
        MsgBox "PORT Error", vbCritical
        On Error Resume Next
        TextPeco.SetFocus
        On Error GoTo 0
    Else
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub txtip_LostFocus()
    Dim sData As String
    If tXTip.Visible = True Then
        sData = Trim(tXTip.TexT)
        If tXTip.TexT <> "" Then
            If Not GeneraTextoIP(sData) Then
                LblInfo2.Caption = CargaCadena(572)  '"Formato IP incorrecto"
                tXTip = ""
            Else
                tXTip.TexT = sData
            End If
        End If
    End If

    '2.0.a error
    On Error Resume Next
    TextPeco.SetFocus
    On Error GoTo 0
    ''''''
End Sub

Private Sub TxtMaquina_KeyPress(KeyAscii As Integer)
    TxtMaquina.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If Len(TxtMaquina.TexT) >= 2 Then TxtMaquina.Locked = True
    If KeyAscii < 48 Or KeyAscii > 59 Then TxtMaquina.Locked = True
End Sub
Public Sub CambiarIdioma()
    FrmMostrador2.Caption = CargaCadena(565)
    CmbEnviar.List(5) = "TouchScale + V12"
    LblInfo.Caption = CargaCadena(524)
    Label2(0).Caption = CargaCadena(554)
    Label2(1).Caption = CargaCadena(54)
    Label2(2).Caption = CargaCadena(552)
    Label1.Caption = CargaCadena(555)
    CmdAceptar.Caption = CargaCadena(550)
    Command1.Caption = CargaCadena(38)
    Command2.Caption = CargaCadena(7)
    Label4.Caption = CargaCadena(48)
End Sub
'1.0.0.4
'Private Sub TxtMaquina_LostFocus()
'    Dim a As Integer
'    Dim MIERROR As Boolean
'    MIERROR = False
'    '*********************************
'    ' comprueba si es necesario, que el número
'    ' de sección máquina sea válido
'    ' para C9/C10
'    '*********************************
'    a = Val(TxtMaquina.TexT)
'    If CmbEnviar.ListIndex = 1 Or CmbEnviar.ListIndex = 5 Then
'    Select Case ModoC9
'        Case 0
'            If (a < 1 Or a > 3) And a <> 6 Then MIERROR = True
'        Case 1
'            If (a < 1 Or a > 5) And a <> 6 Then MIERROR = True
'        Case 2
'            If (a < 1 Or a > 7) And a <> 8 Then MIERROR = True
'    End Select
'    End If
'    If MIERROR Then
'        If IsNumeric(TxtMaquina.TexT) Then
'            If Not UsaGamaAlta And (usaSC10 And TipoC9) Then
'                If MIERROR Then
'                    MsgBox CargaCadena(1019), vbExclamation
'                    TxtMaquina.TexT = ""
'                    TxtMaquina.SetFocus
'                End If
'            End If
'            If UsaGamaAlta And TipoC9 And usaSC10 Then
'                MsgBox CargaCadena(1020), vbInformation
'                CmbEnviar.Enabled = False
'                CmbEnviar.ListIndex = 0
'            End If
'        End If
'    Else
'       If NumeroSistemas > 1 Then CmbEnviar.Enabled = True
'    End If
'    '///cas.n->comprueba que el codigo de seccion para sistemas C9 CODIGO, no sea mayor de uno
'    If (Val(Me.TxtMaquina.TexT) > 1) And (var.C9Cod <> 0) Then
'        MsgBox CargaCadena(1234), vbInformation
'        Me.TxtMaquina.TexT = ""
'        Me.TxtMaquina.SetFocus
'        Exit Sub
'    End If
'End Sub

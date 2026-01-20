VERSION 5.00
Begin VB.Form FrmCambioRapido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Directo a Balanzas"
   ClientHeight    =   6075
   ClientLeft      =   1590
   ClientTop       =   1215
   ClientWidth     =   8190
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtborrar 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   31
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtborrar 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   30
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtborrar 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   29
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtborrar 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   28
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton CmdGoBorrar 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   615
      ItemData        =   "FrmCambioRapido.frx":0000
      Left            =   5040
      List            =   "FrmCambioRapido.frx":0002
      TabIndex        =   26
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar precios por Tramos >>"
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   3600
      Width           =   3255
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1005
      Left            =   4680
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   17
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   14
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   13
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtDato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Precio 2"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tramo 2 (KG)"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Precio 1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tramo 1 (KG)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Nuevo"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PLU"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Mostrador : "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   120
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label4 
      Caption         =   "Se borrará :"
      Height          =   375
      Left            =   5040
      TabIndex        =   27
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Código :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Familia :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Subsección :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3975
   End
End
Attribute VB_Name = "FrmCambioRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'\\\\\\\\\\\\\\
'//////////////
Private Articulo As DB_Articulo
Private sNetPath() As String
Private iNetPath As Integer
'\\\\\\\\\\\\\\
'//////////////
Private Sub Form_Load()
Dim Base As DAO.Database
Dim Registro As DAO.Recordset
Dim b As Integer
    CambiarIdioma
    iNetPath = 0
    Me.Height = 4440
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    Option1(0).Value = True
    cmd(0).Enabled = False
    If Not UsaGamaAlta Then CmdBorrar.Visible = False
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Combo1.AddItem !codi_ident
                .Movenext
            Loop
            If HaySeguridad Then
                If UsuarioActual.Asociado <> 0 Then
                    Combo1.Clear
                    Combo1.AddItem UsuarioActual.Asociado
                    Combo1.TexT = Combo1.List(0)
                    Combo1.Enabled = False
                Else
                    Combo1.TexT = Combo1.List(0)
                End If
            Else
                Combo1.TexT = Combo1.List(0)
            End If
        Else
            For b = 0 To 5
                TxtDato(b).Enabled = False
            Next b
            cmd(0).Enabled = False
            MsgBox CargaCadena(399), vbCritical
        End If
    End With
    Base.Close
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        cmd(0).Caption = "Modificar"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If usaSC10 And Not SC10Rs232 And Not TipoC9 Then
        Envia_Red
    End If
    lCogeTiquet = True
    frmEpelsa.Enabled = True

End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub Cmd_Click(Index As Integer)
Dim Resp As Integer
Dim Resp2 As Boolean
Dim b As Integer
Dim MyPrec As String
Dim Bucle As Long
Dim B2 As Long
Dim NSistema As Long
'********************
Dim ok As Boolean
'********************
    If Index = 0 And Trim(TxtDato(1).TexT) = "" Then TxtDato(1).TexT = Articulo.precio
    If Index = 0 And Trim(TxtDato(1).TexT) = "" Then Exit Sub
    If Index = 0 Then
        cmd(0).Enabled = False
        cmd(1).Enabled = False
        CmdBorrar.Enabled = False
        CmdGoBorrar.Enabled = False
        '******************************************************errores
        If RechazarPlu0 Then
            If ToDouble(TxtDato(1).TexT) = 0 Then
                MsgBox CargaCadena(368), vbCritical
                Exit Sub
            End If
        End If
        If Trim(TxtDato(1).TexT) = "" Then
            MsgBox CargaCadena(1029), vbCritical
            Exit Sub
        Else
            Articulo.precio = ToDouble(TxtDato(1).TexT)
        End If
        If Trim(TxtDato(2).TexT) <> "" Then
            Articulo.TRM1 = ToDouble(TxtDato(2).TexT)
            Articulo.PRC1 = ToDouble(TxtDato(3).TexT)
        Else
            Articulo.TRM1 = -1
            Articulo.PRC1 = -1
        End If
        If Trim(TxtDato(4).TexT) <> "" Then
            Articulo.TRM2 = ToDouble(TxtDato(4).TexT)
            Articulo.PRC2 = ToDouble(TxtDato(5).TexT)
        Else
            Articulo.TRM2 = -1
            Articulo.PRC2 = -1
        End If
        
        Resp = Alta_Articulo(Articulo)
        
        If Dir(App.Path & "\transpluasc.exe") = "" Then
        
        Select Case Resp
            Case 1
                Lista.AddItem CargaCadena(1024)
                If UsaGamaAlta Then
                    NSistema = -1
                    MUE_Sistemas
                    If SisEur.NSistemas > 0 Then
                        For Bucle = 0 To SisEur.NSistemas - 1
                            If SisEur.Sistemas(Bucle).NSecciones > 0 Then
                                For B2 = 0 To SisEur.Sistemas(Bucle).NSecciones - 1
                                    If SisEur.Sistemas(Bucle).Secciones(B2) = Articulo.Mostrador Then
                                        NSistema = Bucle
                                        Exit For
                                    End If
                                Next B2
                            End If
                            If NSistema > -1 Then Exit For
                        Next Bucle
                   End If
                   If NSistema >= 0 Then
                        Lista.AddItem CargaCadena(516) & " -> Euroscale"
                        ok = MUE_EnviaArticuloNW(CInt(NSistema), False, True, Articulo.codigo)
                        '*****************************************************************************************
                        'cas.v151->se corrige el error del no envio de las altas y bajas de precios por tramo
                        'If ok Then MUE_EnviaTramos CInt(NSistema), False, True, Articulo.codigo
                        '********************************************************************************
                        Select Case ok
                            Case True
                                Lista.AddItem "OK"
                            Case Else
                                Lista.AddItem CargaCadena(1025)
                        End Select
                    Else
                        If Articulo.Balenv Then
                            If SisEur.NSistemas > 0 Then
                                Lista.AddItem CargaCadena(516) & " -> Euroscale"
                                For Bucle = 0 To SisEur.NSistemas - 1
                                    Select Case MUE_EnviaArticuloNW(CInt(Bucle), False, True, Articulo.codigo)
                                        Case True
                                            Lista.AddItem "OK"
                                        Case Else
                                            Lista.AddItem CargaCadena(1025)
                                    End Select
                                Next Bucle
                            End If
                        End If
                    End If
                End If
            Case 4
                Lista.AddItem CargaCadena(1026)
            Case 17
                Lista.AddItem CargaCadena(897)
            Case Else
                Lista.AddItem CargaCadena(970)
        End Select
        
        End If
        
        For b = 0 To 5
            TxtDato(b).TexT = ""
        Next b
        cmd(1).Enabled = True
        CmdBorrar.Enabled = True
        TxtDato(0).SetFocus
    Else
        Unload Me
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub CmdBorrar_Click()
    If Me.Height = 6450 Then
        Me.Height = 4440
        CmdBorrar.Caption = CargaCadena(917) & " >>"  '"Borrar precios por Tramos >>"
    Else
        Me.Height = 6450
        CmdBorrar.Caption = CargaCadena(917) & " <<"  ' "Borrar precios por Tramos <<"
        txtborrar(0).TexT = ""
        txtborrar(1).TexT = ""
        txtborrar(2).TexT = ""
        txtborrar(3).TexT = ""
        CmdGoBorrar.Enabled = False
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub CmdGoBorrar_Click()
Dim Resp As Integer
Dim cnt As Long
    Resp = MsgBox(CargaCadena(994), vbExclamation + vbYesNo + vbDefaultButton2)
    If Resp = vbNo Then Exit Sub
    If List1.ListCount > 0 Then
        CmdGoBorrar.Enabled = False
        CmdBorrar.Enabled = False
        cmd(0).Enabled = False
        For cnt = 0 To List1.ListCount - 1
            TxtDato(0).TexT = Val(left(List1.List(cnt), 6))
            PonDato
            TxtDato(2).TexT = ""
            TxtDato(3).TexT = ""
            TxtDato(4).TexT = ""
            TxtDato(5).TexT = ""
            Cmd_Click (0)
        Next cnt
        CmdGoBorrar.Enabled = True
        CmdBorrar.Enabled = True
        cmd(0).Enabled = True
        txtborrar_LostFocus (0)
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub CambiarIdioma()
    CmdGoBorrar.Caption = CargaCadena(38)
    Option1(0).Caption = CargaCadena(69)
    Label1(0).Caption = CargaCadena(69)
    cmd(0).Caption = CargaCadena(43)
    cmd(1).Caption = CargaCadena(7)
    Label2.Caption = CargaCadena(302)
    Label1(1).Caption = CargaCadena(92)
    Label1(2).Caption = CargaCadena(651) & " 1(KG)"
    Label1(4).Caption = CargaCadena(651) & " 2(KG)"
    Label1(3).Caption = CargaCadena(71) & " 1"
    Label1(5).Caption = CargaCadena(71) & " 2"
    Label3(0).Caption = CargaCadena(302)
    Label3(1).Caption = CargaCadena(10)
    Label3(2).Caption = CargaCadena(67)
    Label3(3).Caption = CargaCadena(69)
    Me.Caption = CargaCadena(916)
    CmdBorrar.Caption = CargaCadena(917) & " >>"  '"Borrar precios por Tramos >>"
    Label4.Caption = CargaCadena(918)
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub Envia_Red()
    Dim Arch As Integer
    Dim Bucle As Integer
    If iNetPath > 0 Then
        Arch = FreeFile()
        Open App.Path & "\orden.dat" For Output As #Arch
        For Bucle = 0 To iNetPath - 1
            Print #Arch, sNetPath(Bucle)
        Next Bucle
        iNetPath = 0
        Close #Arch
        MostrarMonitor
        'FrmSC10_EnviaDirectorioRedSC10
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub Option1_Click(Index As Integer)
    Dim b As Integer
    For b = 0 To 5
        TxtDato(b).TexT = ""
    Next b
    Select Case Index
        Case 0
            TxtDato(0).MaxLength = 6
            Label1(0).Caption = CargaCadena(89)
            Combo1.Enabled = False
            Label2.Enabled = False
        Case Else
            TxtDato(0).MaxLength = 4
            Label1(0).Caption = "PLU"
            Label2.Enabled = True
            If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then Combo1.Enabled = True
    End Select
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub txtborrar_KeyPress(Index As Integer, KeyAscii As Integer)
    txtborrar(Index).Locked = Checktexto(KeyAscii, 12, txtborrar(Index).TexT, True, False)
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub txtborrar_LostFocus(Index As Integer)
    Dim Base As DAO.Database
    Dim R1 As DAO.Recordset
    Dim Filtro As String
    List1.Clear
    If Trim(txtborrar(0).TexT) <> "" Or Trim(txtborrar(1).TexT) <> "" Or Trim(txtborrar(2).TexT) <> "" Or Trim(txtborrar(3).TexT) <> "" Then
        Filtro = creamefiltro
        If Filtro <> "" Then
            Set Base = OpenDatabase(Base_General)
            Set R1 = Base.OpenRecordset("select * from articulo where usatramos=true " & Filtro)
            If R1.EOF Then
                List1.AddItem CargaCadena(1027)
                CmdGoBorrar.Enabled = False
                Label4.Caption = CargaCadena(1021) & " : "
            Else
                R1.MoveFirst
                Do Until R1.EOF
                    List1.AddItem Format(R1!codigo, "000000") & "  " & R1!des_plu1
                    R1.Movenext
                Loop
                CmdGoBorrar.Enabled = True
                Label4.Caption = CargaCadena(1021) & " : " & List1.ListCount & " " & CargaCadena(1028)
            End If
            Base.Close
        Else
            CmdGoBorrar.Enabled = False
            Label4.Caption = CargaCadena(1021) & " : "
            List1.AddItem CargaCadena(1022)
        End If
    Else
        Label4.Caption = CargaCadena(1021) & " : "
        List1.AddItem CargaCadena(1022)
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Function creamefiltro() As String
    Dim Filtro As String
    Filtro = ""
    If Trim(txtborrar(0).TexT) <> "" Then
        Filtro = " codi_ident=" & Val(txtborrar(0).TexT)
    End If
    If Trim(txtborrar(1).TexT) <> "" Then
        If Trim(Filtro) <> "" Then
            If Right(Filtro, 3) <> "and" Then Filtro = Filtro & " and"
        End If
        Filtro = Filtro & " codi_sub=" & Val(txtborrar(1).TexT)
    End If
    If Trim(txtborrar(2).TexT) <> "" Then
        If Trim(Filtro) <> "" Then
            If Right(Filtro, 3) <> "and" Then Filtro = Filtro & " and"
        End If
        Filtro = Filtro & " codi_fam=" & Val(txtborrar(2).TexT)
    End If
    If Trim(txtborrar(3).TexT) <> "" Then
        If Trim(Filtro) <> "" Then
            If Right(Filtro, 3) <> "and" Then Filtro = Filtro & " and"
        End If
        Filtro = Filtro & " codigo=" & Val(txtborrar(3).TexT)
    End If
    If Trim(Filtro) <> "" Then
        Filtro = "and " & Filtro
    End If
    creamefiltro = Filtro
End Function
'\\\\\\\\\\\\\\
'//////////////
Private Sub TxtDato_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Max As Integer
    Dim MyDec As Boolean
    MyDec = False
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 1 Then
            If TxtDato(2).TexT = "" Then
                If cmd(0).Enabled Then
                    cmd(0).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If Index < 5 Then
            TxtDato(Index + 1).SetFocus
        Else
            If cmd(0).Enabled Then cmd(0).SetFocus
        End If
    Else
        Select Case Index
            Case 0
                If Option1(0).Value Then
                    Max = 6
                Else
                    Max = 4
                End If
            Case 1
                If UsaEuro Or (Not UsaEuro And decimales > 0) Then MyDec = True
                Max = 7
            Case 2
                If Articulo.WGH = True Then 'no pesado
                    MyDec = False
                    Max = 3
                Else
                    MyDec = True
                    Max = 6
                End If
                'MyDec = True
                'Max = 9
            Case 3
                If UsaEuro Or (Not UsaEuro And decimales > 0) Then MyDec = True
                Max = 7
            Case 4
                If Articulo.WGH = True Then 'no pesado
                    MyDec = False
                    Max = 3
                Else
                    MyDec = True
                    Max = 6
                End If
                'MyDec = True
                'Max = 9
            Case 5
                If UsaEuro Or (Not UsaEuro And decimales > 0) Then MyDec = True
                Max = 7
        End Select
        TxtDato(Index).Locked = Checktexto(KeyAscii, Max, TxtDato(Index).TexT, True, MyDec)
        If (Index = 2 Or Index = 4) And Articulo.WGH = False Then
             
             If Val(TxtDato(Index).TexT & Chr(KeyAscii)) > 30 Then
                 TxtDato(Index).TexT = ""
                 TxtDato(Index + 1).TexT = ""
                 KeyAscii = 0
             End If
             
        End If
    End If
End Sub
'*********
Private Sub TxtDato_Change(Index As Integer)
    If (Index = 2) Or (Index = 3) Then
        If (TxtDato(2).TexT = "") Or (TxtDato(3).TexT = "") Then
            TxtDato(4).TexT = ""
            TxtDato(5).TexT = ""
        End If
    End If
End Sub
'*********
Private Sub TxtDato_LostFocus(Index As Integer)
Dim b As Integer
    If Index = 0 Then
        Lista.Clear
        If Trim(TxtDato(0).TexT) = "" Then
            For b = 1 To 5
                TxtDato(1).TexT = ""
                cmd(0).Enabled = False
            Next b
        Else
            PonDato
        End If
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Public Sub PonDato()
    Dim Data As String
    Dim b As Integer
    Articulo.codigo = 0
    Articulo.plu = 0
    Articulo.Mostrador = 0
    TxtDato(1).TexT = ""
    TxtDato(2).TexT = "0"
    TxtDato(3).TexT = ""
    TxtDato(4).TexT = "0"
    TxtDato(5).TexT = ""
    Label1(6).Caption = ""
    Select Case Option1(0).Value
        Case True
            Articulo.codigo = Val(TxtDato(0).TexT)
        Case False
            Articulo.plu = Val(TxtDato(0).TexT)
            Articulo.Mostrador = Val(Combo1.TexT)
    End Select
    If DB_Consulta_Art(Articulo) Then
        Label1(6).Caption = CargaCadena(658) & " : " & ToString(Articulo.precio)
        If Articulo.TRM1 >= 0 Then
            TxtDato(2).TexT = ToString(Articulo.TRM1)
            TxtDato(3).TexT = ToString(Articulo.PRC1)
        End If
        If Articulo.TRM2 >= 0 Then
            TxtDato(4).TexT = ToString(Articulo.TRM2)
            TxtDato(5).TexT = ToString(Articulo.PRC2)
        End If
        cmd(0).Enabled = True
        If Articulo.WGH Then
            Label1(2) = Mid(Label1(2).Caption, 1, Len(Label1(2).Caption) - 4) & "(Un)"
            Label1(4) = Mid(Label1(4).Caption, 1, Len(Label1(4).Caption) - 4) & "(Un)"
        Else
            Label1(2) = Mid(Label1(2).Caption, 1, Len(Label1(2).Caption) - 4) & "(KG)"
            Label1(4) = Mid(Label1(4).Caption, 1, Len(Label1(4).Caption) - 4) & "(KG)"
        End If
    Else
        cmd(0).Enabled = False
        For b = 0 To 5
            TxtDato(b).TexT = ""
        Next b
        Lista.AddItem CargaCadena(1023)
        TxtDato(0).SetFocus
    End If
End Sub

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Totales"
   ClientHeight    =   4245
   ClientLeft      =   1050
   ClientTop       =   1500
   ClientWidth     =   5430
   Icon            =   "FrmExportar.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox FileTQ 
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Histórico (dd/mm/yy)"
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox TxtNombre 
      Height          =   372
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   5172
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "&Exportar"
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
      TabIndex        =   16
      Top             =   3600
      Width           =   1932
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
      Left            =   3360
      TabIndex        =   15
      Top             =   3600
      Width           =   1932
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   492
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   5172
      _ExtentX        =   9128
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   11
      Left            =   4080
      TabIndex        =   13
      Top             =   1560
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   10
      Left            =   2760
      TabIndex        =   12
      Top             =   1560
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   9
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1212
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1440
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   7
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   6
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   3
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Pulse Sobre el Fichero que desea Exportar"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "FrmExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const BALSAM = False
Private MiTienda As String
Public MostrarMensajes As Boolean
Private SabLineas As Integer
Private CancelarExport As Boolean
Private EstadoExport As Boolean
Private Type TipoFragmento
    id As Integer
    inicio As Long
    fin As Long
    Adicional As String
End Type
Private Fragmento() As TipoFragmento
Private fragmentos As Long
Private GA As String
Private SC10 As String
Private TipoExport As Integer
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        txtfecha.TexT = AdaptaFecha(Format(Now, "dd/mm/yy"))
        txtfecha.Enabled = True
    Else
        txtfecha.Enabled = False
    End If
End Sub

Private Sub CmdCancelar_Click()
    If EstadoExport = False Then
        Unload Me
    End If
End Sub

Public Sub CmdExportar_Click()
    Dim bucle As Integer
    Dim MiFecha As Date
    If Check1.Value = vbChecked Then
        MiFecha = DateSerial(Right(txtfecha.TexT, 2), Mid(txtfecha.TexT, 4, 2), left(txtfecha.TexT, 2))
    End If

    If Trim(TxtNombre.TexT) = "" Then Exit Sub
    For bucle = 0 To 11
        If StrConv(TxtNombre.TexT, vbUpperCase) = StrConv(Command1(bucle).Caption, vbUpperCase) Then
            EstadoExport = True
            If Check1.Value = vbChecked Then
                ExportarFichero bucle, , MiFecha
            Else
                ExportarFichero bucle
            End If
            Exit Sub
        End If
    Next bucle
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim MiFecha As Date
    EstadoExport = True
    If Check1.Value = vbChecked Then
        MiFecha = DateSerial(Right(txtfecha.TexT, 2), Mid(txtfecha.TexT, 4, 2), left(txtfecha.TexT, 2))
    End If
    If MiFecha <> 0 Then
        ExportarFichero Index, , MiFecha
    Else
        ExportarFichero Index
    End If
End Sub

Private Sub Form_Load()
Dim bucle As Integer

    '1.7.2
    'If Not lEsExport Then
    '    Exit Sub
    'End If
    '''''''''''''''''''''''
    
    Me.Caption = db.CargaCadena(1254)
    Me.LblInfo.Caption = db.CargaCadena(1253)
    Me.CmdExportar.Caption = db.CargaCadena(692)
    Me.CmdCancelar.Caption = db.CargaCadena(288)
    Me.Icon = Form2.Icon
    MostrarMensajes = True
    CancelarExport = False
    EstadoExport = False
    For bucle = 0 To 11
        Command1(bucle).Enabled = False
        Command1(bucle).Caption = ""
    Next bucle
    File1.Path = Miruta & "\etc\"
    File1.FileName = "*.fet"
    File1.Refresh
    For bucle = 0 To File1.ListCount - 1
        If bucle = 12 Then Exit For
        Command1(bucle).Enabled = True
        Command1(bucle).Caption = left(File1.List(bucle), Len(File1.List(bucle)) - 4)
    Next bucle
    If Exp_Diario Then
        Check1.Visible = True
        txtfecha.Enabled = False
    Else
        Check1.Visible = False
        txtfecha.Visible = False
    End If
    If HayComandos Then CmdCancelar.Enabled = False
End Sub
Public Sub ExportarFichero(ElBoton As Integer, Optional MiArchivo As String, Optional MiFecha As Date, Optional Continua As Boolean)
    Dim Fichero As Integer
    Dim Buffer As String
    Dim FaEx As String
    If Dir(App.Path & "\noexportados.txt") <> "" Then
        Continua = True
        CadenadeLog "Se exportarán Tiques y líneas no exportadas..."
    End If
    
    MiTienda = ""
    Fichero = FreeFile
    If MiArchivo <> "" Then
        FaEx = MiArchivo
    Else
        FaEx = Command1(ElBoton).Caption
    End If
    Open Miruta & "\etc\" & FaEx & ".fet" For Input As Fichero
    Line Input #Fichero, Buffer
    Close #Fichero
    If InStr(1, Buffer, "+") Then
        If HayMulti Then
            If TiendaActual < 100 Then
                MiTienda = Format(TiendaActual, "00")
            Else
                MiTienda = Format(TiendaActual, "000")
            End If
        End If
        Buffer = Mid(Buffer, InStr(1, Buffer, "+") - 1)
    End If
    If Not IsNumeric(Buffer) Then
        If MostrarMensajes Then MsgBox CargaCadena(848), vbCritical, vbOKOnly
        Exit Sub
    End If
    If MiFecha = 0 Then
        Select Case Val(Buffer)
            Case 1
                'If Continua = False Then
                ExportaGTarti Exp_Path & "\" & FaEx
            Case 2
                'If Continua = False Then
                ExportaGTSecs Exp_Path & "\" & FaEx
            Case 3
                'If Continua = False Then
                ExportaGTVend Exp_Path & "\" & FaEx
            Case 4
                '2.0. 24 If Continua = False Then ExportaCabecera MiruTa & "\" & FaEx
                ExportaCabecera Exp_Path & "\" & FaEx, , Continua
            Case 5
                '2.0.24 If Continua = False Then ExportaLinea MiruTa & "\" & FaEx
                ExportaLinea Exp_Path & "\" & FaEx, , Continua
            Case 6
                ExportaCabLin Exp_Path & "\" & FaEx, , Continua
            Case 7
                ExportaGThora Exp_Path & "\" & FaEx
            Case Else
                If MostrarMensajes Then MsgBox CargaCadena(848), vbCritical, vbOKOnly
                Exit Sub
        End Select
    Else
        Select Case Val(Buffer)
            Case 1
                ExportaGTarti Exp_Path & "\" & FaEx, MiFecha
            Case 2
                ExportaGTSecs Exp_Path & "\" & FaEx, MiFecha
            Case 3
                ExportaGTVend Exp_Path & "\" & FaEx, MiFecha
            Case 4
                ExportaCabecera Exp_Path & "\" & FaEx, MiFecha, Continua
            Case 5
                ExportaLinea Exp_Path & "\" & FaEx, MiFecha, Continua
            Case 6
                ExportaCabLin Exp_Path & "\" & FaEx, MiFecha, Continua
            Case 7
                ExportaGThora Exp_Path & "\" & FaEx, MiFecha
                
            Case Else
                If MostrarMensajes Then MsgBox CargaCadena(848), vbCritical, vbOKOnly
                Exit Sub
        End Select
    End If
End Sub
Private Sub ExportaGTarti(Elfichero As String, Optional MiFecha As Date, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Archivo As Integer
    Dim bucle As Long
    Dim Buffer As String
    Dim Contbuffer As Long
    Dim Cadenilla As String
    Dim laLongitud As Integer
    Dim Sabado_y_Currando As String
    Dim Condicion As String
    Dim nFichHys As Integer
    Dim sBas As String
    Dim lGT As Boolean
    
    If InStr(1, Elfichero, "fin_dbal") <> 0 Then lGT = True
    If MiFecha = 0 Then
        MiFecha = Now
    End If
    'enzo
    'If Exp_Diario = True Then
    '    'Condicion = " where fecha=#" & Format(MiFecha, "mm/dd/yyyy") & "#"
    '    Condicion = " where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    'Else
    '    Condicion = ""
    'End If
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            'If (lCont) Then
            '    Condicion = " where  ( (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            'Else
                Condicion = " where cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")"
            'End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    LeeFragmentos (Elfichero)
    'Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from gtarti" & Condicion & " order by cdate(fecha)")
    With Registro
        If Not .EOF Then
            .MoveFirst
            .MoveLast
            Barra.Max = .Recordcount + 1
            Barra.Value = 0
            Archivo = FreeFile
            On Error Resume Next
            Open Exp_CambiaPath(Elfichero) For Output As #Archivo
            If Err.Number <> 0 Then
                MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                Exit Sub
            End If
            On Error GoTo 0
            .MoveFirst
            Do Until .EOF
                Barra.Value = Barra.Value + 1
                Buffer = ""
                Contbuffer = 1
                For bucle = 1 To fragmentos
                    Do While Contbuffer < Fragmento(bucle).inicio
                        Contbuffer = Contbuffer + 1
                        Buffer = Buffer & " "
                    Loop
                    laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
                    Select Case Val(Fragmento(bucle).id)
                        Case 0
                            Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                            If lGT And Val(Fragmento(bucle).Adicional) = 32 Then
                                Buffer = Buffer & Format(nGTOrd, "0000")
                                laLongitud = laLongitud + 4
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 1
                            Sabado_y_Currando = FormFecha(Format(!Fecha, "dd/mm/yy"))
                            Buffer = Buffer & Formatea(Sabado_y_Currando, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 2
                            Buffer = Buffer & Formatea(!Plu, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 3
                            Buffer = Buffer & Formatea(!des_plu1, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 4
                            Buffer = Buffer & Formatea(!codi_ident, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 5
                            Buffer = Buffer & Formatea(!codi_sub, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 6
                            Buffer = Buffer & Formatea(!codi_fam, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 7
                            Buffer = Buffer & Formatea(!codigo, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 8
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!precio * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!precio * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 9
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!pesetas * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!pesetas * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 10
                            'If Not Sabeco Then
                                Buffer = Buffer & Formatea(!Unidades, laLongitud, True)
                                Contbuffer = Contbuffer + laLongitud
                            'Else
                            '    Buffer = Buffer & Formatea(!Unidades * 1000, laLongitud, True)
                            '    Contbuffer = Contbuffer + laLongitud
                            'End If
                        Case 11
                            Buffer = Buffer & Formatea(1000 * !Peso, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 12
                            If !Unidades = 0 Then
                                Buffer = Buffer & Formatea(1000 * !Peso, laLongitud, True)
                            Else
                                'If Not Sabeco Then
                                    Buffer = Buffer & Formatea(!Unidades, laLongitud, True)
                                'Else
                                '    Buffer = Buffer & Formatea(!Unidades * 1000, laLongitud, True)
                                'End If
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 13
                            Buffer = Buffer & Formatea(!tara, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 14
                            Buffer = Buffer & Formatea(!caducidad, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 15
                            Buffer = Buffer & Formatea(!operacion, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 16
                            Buffer = Buffer & Formatea(!tipo_iva, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 17
                            If !Modificado = False Then
                                Buffer = Buffer & Formatea("0", laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea("1", laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 18
                            Buffer = Buffer & Formatea(!secc_maqui, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 19
                            If !pesetas >= 0 Then
                                Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                            Else
                                Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 20
                            If !Unidades <> 0 Then
                                Buffer = Buffer & Formatea("2", laLongitud, False)
                            Else
                                If !Peso <> 0 Then
                                    Buffer = Buffer & Formatea("1", laLongitud, False)
                                Else
                                    Buffer = Buffer & Formatea("0", laLongitud, False)
                                End If
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 21
                            Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                    End Select
                Next bucle
                Print #Archivo, Buffer
                .Movenext
            Loop
            Close #Archivo
            'If MostrarMensajes Then MsgBox CargaCadena( 849)
        Else
            'If MostrarMensajes Then MsgBox CargaCadena( 850)
        End If
    End With
    Unload Me
End Sub
Private Sub ExportaGTSecs(Elfichero As String, Optional MiFecha As Date, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Buffer As String
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim Archivo As Integer
    Dim Contbuffer As Long
    Dim Sabado_y_Currando As String
    Dim Condicion As String
    Dim nFichHys As Integer
    Dim sBas As String
    
    If MiFecha = 0 Then
        MiFecha = Now
    End If
    
    'If Exp_Diario = True Then
    '    'Condicion = " where fecha=#" & Format(MiFecha, "mm/dd/yyyy") & "#"
    '    Condicion = " where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    'Else
    '    Condicion = ""
    'End If
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            'If (lCont) Then
            '    Condicion = " where  ( (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            'Else
                Condicion = " where cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")"
            'End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    LeeFragmentos (Elfichero)
    'Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from gtsecs" & Condicion & " order by cdate(fecha)")
    With Registro
        If Not .EOF Then
            Archivo = FreeFile
            On Error Resume Next
            Open Exp_CambiaPath(Elfichero) For Output As #Archivo
            If Err.Number <> 0 Then
                MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                Exit Sub
            End If
            On Error GoTo 0
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                Contbuffer = 1
                For bucle = 1 To fragmentos
                    Do While Contbuffer < Fragmento(bucle).inicio
                        Contbuffer = Contbuffer + 1
                        Buffer = Buffer & " "
                    Loop
                    laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
                    Select Case Val(Fragmento(bucle).id)
                        Case 0
                            Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 1
                            Sabado_y_Currando = FormFecha(Format(!Fecha, "dd/mm/yy"))
                            Buffer = Buffer & Formatea(Sabado_y_Currando, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 2
                            Buffer = Buffer & Formatea(!codi_ident, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 3
                            Buffer = Buffer & Formatea(!codi_sub, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 4
                            Buffer = Buffer & Formatea(!codi_fam, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 5
                            Buffer = Buffer & Formatea(!tiquets, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 6
                            Buffer = Buffer & Formatea(!operacion, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 7
                            Buffer = Buffer & Formatea(!Peso, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 8
                            Buffer = Buffer & Formatea(!Unidades, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 9
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!pesetas * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!pesetas * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 10
                            Buffer = Buffer & Formatea(!secc_maqui, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 11
                            If !Importe >= 0 Then
                                Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                            Else
                                Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 12
                            Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                    End Select
                Next bucle
                Print #Archivo, Buffer
                .Movenext
            Loop
            Close #Archivo
                'If MostrarMensajes Then MsgBox CargaCadena( 849)
            Else
                'If MostrarMensajes Then MsgBox CargaCadena( 850)
        End If
    End With
End Sub

Private Sub ExportaGThora(Elfichero As String, Optional MiFecha As Date, Optional ByVal mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Buffer As String
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim Archivo As Integer
    Dim Contbuffer As Long
    Dim Sabado_y_Currando As String
    Dim Condicion As String
    Dim sPar As String
    Dim nFichHys As Integer
    Dim sBas As String
    
    
    If MiFecha = 0 Then
        MiFecha = Now
    End If
    sPar = CDate(MiFecha)
    If Len(sPar) > 10 Then sPar = Mid(sPar, 1, 10)
    
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            'If (lCont) Then
            '    Condicion = " where  ( (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            'Else
                Condicion = " where cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")"
            'End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    LeeFragmentos (Elfichero)
    '
    Set Registro = Base.OpenRecordset("select * from gthora" & Condicion & " order by cdate(fecha)")
    With Registro
        If Not .EOF Then
            Archivo = FreeFile
            On Error Resume Next
            Open Exp_CambiaPath(Elfichero) For Output As #Archivo
            If Err.Number <> 0 Then
                MsgBox "Error. File or Folder not Available : " & Exp_CambiaPath(Elfichero), vbCritical
                Exit Sub
            End If
            On Error GoTo 0
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                Contbuffer = 1
                For bucle = 1 To fragmentos
                    Do While Contbuffer < Fragmento(bucle).inicio
                        Contbuffer = Contbuffer + 1
                        Buffer = Buffer & " "
                    Loop
                    laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
                    Select Case Val(Fragmento(bucle).id)
                        Case 0
                            Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 1
                            Sabado_y_Currando = FormFecha(Format(!Fecha, "dd/mm/yy"))
                            Buffer = Buffer & Formatea(Sabado_y_Currando, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 2
                            Buffer = Buffer & Formatea(FormHora(Format(!hora, "hh:mm")), laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 3
                            Buffer = Buffer & Formatea(!tiquet, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 4
                            Buffer = Buffer & Formatea(!operaciones, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 5
                            Buffer = Buffer & Formatea(!operaciones_can, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 6
                            Buffer = Buffer & Formatea(!operaciones_neg, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 7
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!Importe * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!Importe * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 8
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!importe_can * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!importe_can * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 9
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!importe_neg * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!importe_neg * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 10
                            Buffer = Buffer & Formatea(!Peso * 1000, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        
                        Case 11
                            Buffer = Buffer & Formatea(!Peso_can * 1000, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 12
                            Buffer = Buffer & Formatea(!Peso_neg * 1000, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 13
                            Buffer = Buffer & Formatea(!Unidades, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 14
                            Buffer = Buffer & Formatea(!unidades_can, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 15
                            Buffer = Buffer & Formatea(!unidades_neg, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                    End Select
                Next bucle
                If Dir(App.Path & "\hucafa.txt") <> "" Then
                    Print #Archivo, Format(TiendaActual, "00") & Buffer
                Else
                    Print #Archivo, Buffer
                End If

                .Movenext
            Loop
            Close #Archivo
                'If MostrarMensajes Then MsgBox CargaCadena( 849)
            Else
                'If MostrarMensajes Then MsgBox CargaCadena( 850)
        End If
    End With
End Sub

Private Sub ExportaGTVend(Elfichero As String, Optional MiFecha As Date, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Archivo As Integer
    Dim Buffer As String
    Dim Contbuffer As Long
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim Sabado_y_Currando As String
    Dim Condicion As String
    'If MiFecha = 0 Then
    '    MiFecha = Now
    'End If
    'If Exp_Diario = True Then
    '    'Condicion = " where fecha=#" & Format(MiFecha, "mm/dd/yyyy") & "#"
    '    Condicion = " where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    'Else
    '    Condicion = ""
    'End If
    Dim nFichHys As Integer
    Dim sBas As String
    
    If MiFecha = 0 Then
        MiFecha = Now
    End If
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            'If (lCont) Then
            '    Condicion = " where  ( (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            'Else
                Condicion = " where cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")"
            'End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    LeeFragmentos (Elfichero)
    'Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from gtvend" & Condicion & " order by cdate(fecha)")
    With Registro
        If Not .EOF Then
            Archivo = FreeFile
            On Error Resume Next
            Open Exp_CambiaPath(Elfichero) For Output As #Archivo
            If Err.Number <> 0 Then
                MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                Exit Sub
            End If
            On Error GoTo 0
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                Contbuffer = 1
                For bucle = 1 To fragmentos
                    Do While Contbuffer < Fragmento(bucle).inicio
                        Contbuffer = Contbuffer + 1
                        Buffer = Buffer & " "
                    Loop
                    laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
                    Select Case Val(Fragmento(bucle).id)
                        Case 0
                            Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 1
                            Sabado_y_Currando = FormFecha(Format(!Fecha, "dd/mm/yy"))
                            Buffer = Buffer & Formatea(Sabado_y_Currando, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 2
                            Buffer = Buffer & Formatea(!codi_ident, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 3
                            Buffer = Buffer & Formatea(!ident_vend, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 4
                            Buffer = Buffer & Formatea(!Nombre, laLongitud, False)
                            Contbuffer = Contbuffer + laLongitud
                        Case 5
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!Supertot * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!Supertot * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 6
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!totcance * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!totcance * (10 ^ decimales), laLongitud, True)
                            End If
                            
                            Contbuffer = Contbuffer + laLongitud
                        Case 7
                            If UsaEuro Then
                                Buffer = Buffer & Formatea(!Total * 100, laLongitud, True)
                            Else
                                Buffer = Buffer & Formatea(!Total * (10 ^ decimales), laLongitud, True)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 8
                            Buffer = Buffer & Formatea(!tiquets, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 9
                            Buffer = Buffer & Formatea(!operacion, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 10
                            Buffer = Buffer & Formatea(!Peso * 1000, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 11
                            Buffer = Buffer & Formatea(!unidad, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 12
                            Buffer = Buffer & Formatea(!cancela, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 13
                            Buffer = Buffer & Formatea(!secc_maqui, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                        Case 14
                            If !Total >= 0 Then
                                Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                            Else
                                Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                            End If
                            Contbuffer = Contbuffer + laLongitud
                        Case 15
                            Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                            Contbuffer = Contbuffer + laLongitud
                    End Select
                Next bucle
                Print #Archivo, Buffer
                .Movenext
                 
            Loop
            
            Close #Archivo
                'If MostrarMensajes Then MsgBox CargaCadena( 849)
        Else
                'If MostrarMensajes Then MsgBox CargaCadena( 850)
        End If
    End With
End Sub
Private Sub ExportaCabecera(Elfichero As String, Optional MiFecha As Date, Optional lCont As Boolean, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Archivo As Integer
    Dim bucle As Long
    Dim Buffer As String
    Dim Contbuffer As Long
    Dim Cadenilla As String
    Dim laLongitud As Integer
    Dim CMax As Integer
    Dim CMin As Integer
    Dim BucleGrande As Long
    Dim Condicion As String
    Dim NSistema As Integer 'c2f/caspiunza
    Dim nSisLoop As Integer 'c2f/caspiunza
    Dim nFichHys As Integer
    Dim sBas As String
    Dim lEnExport As Boolean
    Dim lCsv As Boolean
    Dim sMicabCsv As String
    'c2f/caspiunza
    NSistema = 1
    lCsv = False
    sMicabCsv = ""
    If InStr(1, Elfichero, ".csv") <> 0 Then
        'lCsv = True
        If Dir(App.Path & "\cabcabcsv.cfg") <> "" Then
            CMin = FreeFile()
            Open App.Path & "\cabcabcsv.cfg" For Input As #CMin
            Line Input #CMin, sMicabCsv
            Close #CMin
            CMin = 0
            lCsv = True
        
        End If
    End If
    ''''''''''''''
    'If MiFecha = 0 Then
    '    MiFecha = Now
    'End If
    'If Exp_Diario = True Then
    '    If (lCont) Then
    '        Condicion = " where ( (cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & ") or (reserv3 <>'SI') )"
    '    Else
    '        Condicion = " where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    '    End If
    'Else
    '    Condicion = ""
    'End If
    
    If MiFecha = 0 Then
        MiFecha = Now
    End If
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            If (lCont) Then
                Condicion = " where  ( (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            Else
                Condicion = " where cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ") or (reserv3 <>'SI')"
            End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    LeeFragmentos (Elfichero)
    If TipoExport = 1 Then
        CMax = 1
        CMin = 1
    End If
    If TipoExport = 2 Then
        CMax = 2
        CMin = 2
    End If
    If TipoExport = 3 Then
        CMax = 2
        CMin = 1
    End If
    'Set Base = OpenDatabase(Base_General)
    lEnExport = False
    
    'For nSisLoop = 1 To NSistema 'c2f/caspiunza
    
    'For BucleGrande = CMin To CMax
        
        'If nSisLoop = 1 Then 'c2f/caspiunza
            'If BucleGrande = 1 Then                Set Registro = Base.OpenRecordset("select * from cabecera" & Condicion & " order by cdate(fecha)")
                If lCsv Then
                    Set Registro = Base.OpenRecordset("select * from cabecera" & Condicion & " order by left(factura,2),right(factura,6)")
                Else
                    Set Registro = Base.OpenRecordset("select * from cabecera" & Condicion & " order by cdate(fecha)")
                End If
            'Else
            '    Set Registro = Base.OpenRecordset("select * from cabeceraSC10" & Condicion)
            'End If
        'Else
        '        Set Registro = Base.OpenRecordset("select * from cabeceratq" & Condicion)
        'End If
        With Registro
            If Not .EOF Then
                .MoveFirst
                .MoveLast
                Barra.Max = .Recordcount + 1
                Barra.Value = 0
                On Error Resume Next
                If lEnExport = False Then
                    Archivo = FreeFile
                    Open Exp_CambiaPath(Elfichero) For Output As #Archivo
                    If Err.Number <> 0 Then
                        MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                        Exit Sub
                    End If
                    If lCsv And sMicabCsv <> "" Then
                        Print #Archivo, sMicabCsv
                    End If
                    lEnExport = True
                End If
                On Error GoTo 0
                .MoveFirst
                Do Until .EOF
                    Barra.Value = Barra.Value + 1
                    Buffer = ""
                    Contbuffer = 1
                    If lCsv Then
                        ExportaunaCabeceraCSV Buffer, Contbuffer, Registro, BucleGrande, lCont, lCsv '2.0.24 (no estaba ,lCont -->Continua)
                    Else
                        ExportaunaCabecera Buffer, Contbuffer, Registro, BucleGrande, lCont '2.0.24 (no estaba ,lCont -->Continua)
                    End If
                    
                    Print #Archivo, Buffer
                    .Movenext
                Loop
        End If
    End With
    
    'If nSisLoop = 2 Then Exit For 'c2f/caspiunza
    
    'Next BucleGrande
    
    'Next nSisLoop
    
    If lEnExport Then Close #Archivo
    lEnExport = False
    'If MostrarMensajes Then MsgBox CargaCadena( 849)
    Unload Me
End Sub
Private Sub ExportaunaCabeceraCSV(Buffer As String, ByRef Contbuffer As Long, Registro As dao.Recordset, _
          ByRef BucleGrande As Long, Optional Continua As Boolean, Optional lEspecial As Boolean)
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim MiValor As Double
    Dim sM As String
    
    'On Error Resume Next
    'If Not lEspecial Then
    '   Dim lEspecial As Boolean
    '   lEspecial = False
    'End If
    'On Error GoTo 0
    If IsNull(lEspecial) Then lEspecial = False
    
    With Registro
    For bucle = 1 To fragmentos
        Do While Contbuffer < Fragmento(bucle).inicio
            Contbuffer = Contbuffer + 1
            'Buffer = Buffer & " "
        Loop
        laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
        Select Case Val(Fragmento(bucle).id)
            Case 0
                Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 1
                'If TipoExport <> 3 Then
                    Buffer = Buffer & Formatea(Val(!nume), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'Else
                '    If BucleGrande = 1 Then
                '        Buffer = Buffer & GA
                '    Else
                '        Buffer = Buffer & SC10
                '    End If
                '    Buffer = Buffer & Formatea(Val(!nume), laLongitud - 1, True)
                '    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 2
                Buffer = Buffer & Formatea(Val(!seccion), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 3
                Buffer = Buffer & Formatea(!Balanza, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 4
                Buffer = Buffer & Formatea(Val(!vended), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 5
                'If Sabeco Then
                '    Buffer = Buffer & Formatea(Val(!numlin) + Val(!lincan), laLongitud, True)
                '    Contbuffer = Contbuffer + laLongitud
                'Else
                    Buffer = Buffer & Formatea(Val(!numlin), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 6
                Buffer = Buffer & Formatea(Val(!lincan), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 7
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !Importe * 100
                    Else
                        MiValor = !Importe * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!Importe)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 8
                Buffer = Buffer & Formatea(FormHora(!hora), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 9
                'If Not Sabeco Then
                    Buffer = Buffer & Formatea(Val(!STPETIC), laLongitud, True)
                    If lEspecial Then
                        If Val(!STPETIC) = 64 Then
                            Buffer = Buffer & ";" & CStr(!Importe)
                        Else
                            Buffer = Buffer & ";" & "0,00"
                        End If
                    End If
                'Else
                '    If (Val(!STPETIC) And 64) Or (Val(!STPETIC) And 128) Then
                '        Buffer = Buffer & Formatea(1, laLongitud, True)
                '    Else
                '        Buffer = Buffer & Formatea(0, laLongitud, True)
                '    End If
                'End If
                Contbuffer = Contbuffer + laLongitud
            Case 10
                Buffer = Buffer & Formatea(!codcli, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 11
                Buffer = Buffer & Formatea(FormFecha(!Fecha), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 12
                Buffer = Buffer & Formatea(Val(!TYPTIC), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 13
                If Trim(!codbar) <> "" Then
                    Buffer = Buffer & Formatea(Mid(!codbar, Val(left(Fragmento(bucle).Adicional, 2)), Val(Right(Fragmento(bucle).Adicional, 2))), laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Space(13), laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 14
                If !Importe >= 0 Then
                    Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 15
                Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 16 'c2f invicta forma de pago
                If Not IsNull(!tipoPago) Then
                    Buffer = Buffer & Formatea(Val(!tipoPago), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 17 'c2f invicta importe bruto
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !importe_bruto * 100
                    Else
                        MiValor = !importe_bruto * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!importe_bruto)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 18 'c2f lote
                Buffer = Buffer & Formatea(!numlote, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 19 'c2f factura
                On Error Resume Next
                Buffer = Buffer & Formatea(!factura, laLongitud, False)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Buffer = Buffer & Formatea(" ", laLongitud, False)
                End If
                On Error GoTo 0
                Contbuffer = Contbuffer + laLongitud
            Case 20 '% 0
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!por0 * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!por0 * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!por0)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
            
                Contbuffer = Contbuffer + laLongitud
            Case 21 '% 1
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!por1 * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!por1 * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!por1)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
            
                Contbuffer = Contbuffer + laLongitud
            Case 22 '% 2
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!por2 * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!por2 * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!por2)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
                
                Contbuffer = Contbuffer + laLongitud
            Case 23 '% 3
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!por3 * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!por3 * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!por3)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 24 '% 4
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!por4 * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!por4 * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!por4)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
        
            Case 25 'base0
                'If !base0 <> 0 Then
                '    MiValor = MiValor
                'End If
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !base0 * 100
                    Else
                        MiValor = !base0 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!base0)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 26 '1
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !base1 * 100
                    Else
                        MiValor = !base1 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!base1)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 27 '2
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !base2 * 100
                    Else
                        MiValor = !base2 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!base2)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 28 '3
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !base3 * 100
                    Else
                        MiValor = !base3 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!base3)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 29 '4
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !base4 * 100
                    Else
                        MiValor = !base4 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!base4)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
        
            Case 30 'imp0
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !imp0 * 100
                    Else
                        MiValor = !imp0 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!imp0)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 31 '1
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !imp1 * 100
                    Else
                        MiValor = !imp1 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!imp1)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 32 '2
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !imp2 * 100
                    Else
                        MiValor = !imp2 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!imp2)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 33 '3
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !imp3 * 100
                    Else
                        MiValor = !imp3 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!imp3)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 34 '4
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !imp4 * 100
                    Else
                        MiValor = !imp4 * (10 ^ decimales)
                    End If
                End If
                'cazani ver If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!imp4)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    If (!STPETIC) = "64" Then sM = "0,00"
                    'cazani ver sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 35
                
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!descuento * 100, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!descuento * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!descuento)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 36 'T. Units
                Buffer = Buffer & Formatea(!TotalUnits, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 37
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!TotalWeight * 1000, laLongitud, True)
                    'BufferBis = BufferBis & Formatea((!TotalWeight * 1000) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!TotalWeight)
                    If IsNull(sM) Or sM = "0" Then sM = "0,000"
                    sM = formatic(sM, 3, laLongitud)
                    Buffer = Buffer & sM
                    'BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
                
        End Select
    Next bucle
    'If Continua Then
        Edit_Record Registro
        !reserv3 = "SI"
        .Update
    'End If
    End With
End Sub

'''Exportación tique actual...(hydratouch 1.0.0)
Private Sub ExportaLinea(Elfichero As String, Optional MiFecha As Date, Optional lCont As Boolean, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Archivo As Integer
    Dim bucle As Long
    Dim Buffer As String
    Dim BufferBis As String
    Dim Contbuffer As Long
    Dim Cadenilla As String
    Dim CMax As Integer
    Dim CMin As Integer
    Dim BucleGrande As Long
    Dim Condicion As String
    Dim ContReg As Long
    Dim BGrande As Long
    Dim NSistema As Integer 'c2f/caspiunza
    Dim nSisLoop As Integer 'c2f/caspiunza
    Dim sSs As String
    Dim nFichHys As Integer
    Dim sBas As String
    Dim x As String
    Dim lEnExport As Boolean
    
    x = ""
    'c2f/caspiunza
    'NSistema = 2
    ''''''''''''''
    CadenadeLog "Fecha Parámetro:" & CDate(MiFecha)
    'If MiFecha = 0 Then
    '    MiFecha = Now
    '    CadenadeLog "Transformada Fecha Parámetro:" & CStr(MiFecha)
    'End If
    'MiFecha = Format(MiFecha, "dd/mm/yy")
    'If Exp_Diario = True Then
    '    'c2f decepal--> hay que seleccionar los no exportados...
    '    'If Dir(App.Path & "\decepal.txt") <> "" Then 'Taquion And TaquionFichero = "" And sunTicket = "" Then
    '    '    Condicion = " where reserv3=" & Chr(34) & "NO" & Chr(34) & " and cstr(cdate(date))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    '    'Else
    '        If (lCont) Then
    '            Condicion = " where  ( (cstr(cdate(date))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & ") and (reserv3 <>'SI') )"
    '        Else
    '            Condicion = " where cstr(cdate(date))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34)
    '        End If
    '    'End If
    'Else
    '    Condicion = ""
    'End If
    'CadenadeLog "Condición:" & Condicion
    If MiFecha = 0 Then
        MiFecha = Now
        CadenadeLog "Transformada Fecha Parámetro:" & CStr(MiFecha)
    End If
    MiFecha = Format(MiFecha, "dd/mm/yy")
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       'select * from tickets where cdate(date)>=cdate("14/11/2016") and cdate(date)<=cdate("29/11/2016") order by val(nume),val(numlin)
       Condicion = " where cdate(date)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(date)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            If (lCont) Then
                Condicion = " where  ( (cdate(date)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")) and (reserv3 <>'SI') )"
            Else
                Condicion = " where cdate(date)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")"
            End If
        Else
            Condicion = ""
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    LeeFragmentos (Elfichero)
    If TipoExport = 1 Then
        CMax = 1
        CMin = 1
    End If
    If TipoExport = 2 Then
        CMax = 2
        CMin = 2
    End If
    If TipoExport = 3 Then
        CMax = 2
        CMin = 1
    End If
    
    lEnExport = False
    
    'Set Base = OpenDatabase(Base_General)
    
    'For nSisLoop = 1 To NSistema 'c2f/caspiunza
    
    'For BucleGrande = CMin To CMax
        'If nSisLoop = 1 Then
        
            'If BucleGrande = 1 Then
                Set Registro = Base.OpenRecordset("select * from tickets" & Condicion & " order by cdate(date),val(nume),val(numlin)")
            'Else
            '    Set Registro = Base.OpenRecordset("select * from ticketsSC10" & Condicion)
            'End If
        
        'Else
        '    Set Registro = Base.OpenRecordset("select * from ticketstq" & Condicion)
        'End If
        
        With Registro
            If Not .EOF Then
                .MoveFirst
                .MoveLast
                Barra.Max = .Recordcount + 1
                ContReg = .Recordcount
                Barra.Value = 0
                '???
                'Archivo = FreeFile
                'On Error Resume Next
                'sSs = Exp_CambiaPath(Elfichero)
                'Open sSs For Output As #Archivo
                '???
'???
                If lEnExport = False Then
                    Archivo = FreeFile
                    On Error Resume Next
                    Open Exp_CambiaPath(Elfichero) For Output As #Archivo
                    If Err.Number <> 0 Then
                        MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                        Exit Sub
                    End If
                    On Error GoTo 0
                    lEnExport = True
                End If

'???
                
                '???
                'If Err.Number <> 0 Then
                '    MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                '    Exit Sub
                'End If
                '???
                
                .MoveFirst
                For bucle = 1 To ContReg
                    Buffer = ""
                    BufferBis = ""
                    Contbuffer = 1
                    Barra.Value = Barra.Value + 1
                    
                    Call ExportaunaLinea(Buffer, BufferBis, Contbuffer, Registro, BucleGrande, x, lCont, False) '2.0.24
                    
                    Print #Archivo, Buffer
                    If Trim(BufferBis) <> "" Then
                        Print #Archivo, BufferBis
                    End If
                    
                    .Movenext
                Next bucle
        End If
    End With
    
    'If nSisLoop = 2 Then Exit For 'c2f/caspiunza
    
    'Next BucleGrande
    
    'Next nSisLoop
    
    '???Close #Archivo
    If lEnExport Then Close #Archivo
    lEnExport = False
    'If MostrarMensajes Then MsgBox CargaCadena( 849)
    
    Unload Me
End Sub
Public Sub ExportaCabLin(Elfichero As String, Optional MiFecha As Date, Optional Continua As Boolean, Optional mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegLin As dao.Recordset
    Dim Archivo As Integer
    Dim ArchIn As Integer
    Dim bucle As Long
    Dim Buffer As String
    Dim BufferBis As String
    Dim BufCab As String
    Dim Contbuffer As Long
    Dim Cadenilla As String
    Dim laLongitud As Integer
    Dim CMax As Integer
    Dim CMin As Integer
    Dim BucleGrande As Long
    Dim ElFicheroCabecera As String
    Dim ElFicheroLineas As String
    Dim FCab() As TipoFragmento
    Dim Fcabs As Integer
    Dim FLin() As TipoFragmento
    Dim FLins As Integer
    Dim BFrag As Integer
    Dim CadenaEnlace As String
    Dim Condicion As String
    Dim Cond2 As String
    Dim Buf2 As String
    Dim Buf3 As String
    Dim Adicional As String
    Dim CuentaEvents As Integer
    Dim lSeHaCreado As Boolean
    'c2f/caspiunza
    Dim sSs As String
    ''''''''''''''
    Dim nFichHys As Integer
    Dim sBas As String
    
    'If Continua Then
    '    Adicional = " and reserv3=" & Chr(34) & "NO" & Chr(34)
    'End If
    'If MiFecha = 0 Then
    '    MiFecha = Now
    'End If
    'If Exp_Diario = True Then
    '    Condicion = " where (cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & Adicional & ")"
    '    Cond2 = " and (cstr(cdate(date))=" & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & Adicional & ")"
    'Else
    '    If Not Continua Then
    '        Condicion = ""
    '        Cond2 = ""
    '    Else
    '        Condicion = " where reserv3=" & Chr(34) & "NO" & Chr(34)
    '        Cond2 = " and reserv3=" & Chr(34) & "NO" & Chr(34)
    '    End If
    'End If
    'If Elfichero = "bbtttt.dat" Then 'aplicaciones y sistemas
    '        Condicion = " where reserv3=" & Chr(34) & "NO" & Chr(34)
    '        Cond2 = " and reserv3=" & Chr(34) & "NO" & Chr(34)
    'End If
''''
''''


    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sBas
       Close #nFichHys
       Set Base = OpenDatabase(sBas)
       Condicion = " where cdate(fecha)>=cdate(" & Chr(34) & MiFecha & Chr(34) & ") and cdate(fecha)<=cdate(" & Chr(34) & mIfECHA2 & Chr(34) & ")"
    Else
        Set Base = OpenDatabase(Base_General)
        If Continua Then
            Adicional = " and reserv3=" & Chr(34) & "NO" & Chr(34)
        End If
        If MiFecha = 0 Then
            MiFecha = Now
        End If
        MiFecha = Format(MiFecha, "dd/mm/yy")
        If Exp_Diario = True Then
            Condicion = " where (cdate(fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")" & Adicional & ")"
            Cond2 = " and (cdate(date)=cdate(" & Chr(34) & MiFecha & Chr(34) & ")" & Adicional & ")"
        Else
            If Not Continua Then
                Condicion = ""
                Cond2 = ""
            Else
                Condicion = " where reserv3=" & Chr(34) & "NO" & Chr(34)
                Cond2 = " and reserv3=" & Chr(34) & "NO" & Chr(34)
            End If
        End If
        If Elfichero = "bbtttt.dat" Then 'aplicaciones y sistemas
                Condicion = " where reserv3=" & Chr(34) & "NO" & Chr(34)
                Cond2 = " and reserv3=" & Chr(34) & "NO" & Chr(34)
        End If
        Set Base = OpenDatabase(Base_General)
    End If
    
    '*************************
    ' intenta leer ficheros de cabecera y lineas
    '*************************
    Buf2 = ""
    For bucle = Len(Elfichero) To 1 Step -1
        If Mid(Elfichero, bucle, 1) <> "\" Then
            Buf2 = Buf2 & Mid(Elfichero, bucle, 1)
        Else
            Exit For
        End If
    Next bucle
    Buf3 = ""
    For bucle = Len(Buf2) To 1 Step -1
        Buf3 = Buf3 & Mid(Buf2, bucle, 1)
    Next bucle
    ArchIn = FreeFile()
    On Error GoTo errorlectura
    Open Miruta & "\etc\" & Buf3 & ".fet" For Input As #ArchIn
    Line Input #ArchIn, Buffer
    Line Input #ArchIn, Buffer
    ElFicheroCabecera = Trim(Miruta & "\" & Buffer)
    Line Input #ArchIn, Buffer
    ElFicheroLineas = Trim(Miruta & "\" & Buffer)
    Close #ArchIn
    On Error GoTo 0
    If ElFicheroCabecera = "" Or ElFicheroLineas = "" Then GoTo errorlectura
    '***********************************
    ' almacena información de las cabeceras
    '***********************************
    LeeFragmentos (ElFicheroCabecera)
    Fcabs = fragmentos
    ReDim FCab(fragmentos + 1)
    For BFrag = 1 To fragmentos
        FCab(BFrag).Adicional = Fragmento(BFrag).Adicional
        FCab(BFrag).fin = Fragmento(BFrag).fin
        FCab(BFrag).id = Fragmento(BFrag).id
        FCab(BFrag).inicio = Fragmento(BFrag).inicio
    Next BFrag
    If TipoExport = 1 Then
        CMax = 1
        CMin = 1
    End If
    If TipoExport = 2 Then
        CMax = 2
        CMin = 2
    End If
    If TipoExport = 3 Then
        CMax = 2
        CMin = 1
    End If
    '***********************************
    ' almacena formato de las líneas
    '***********************************
    LeeFragmentos (ElFicheroLineas)
    FLins = fragmentos
    ReDim FLin(fragmentos + 1)
    For BFrag = 1 To fragmentos
        FLin(BFrag).Adicional = Fragmento(BFrag).Adicional
        FLin(BFrag).fin = Fragmento(BFrag).fin
        FLin(BFrag).id = Fragmento(BFrag).id
        FLin(BFrag).inicio = Fragmento(BFrag).inicio
    Next BFrag
    '********************************
    ' exporta
    '********************************
    'Set Base = OpenDatabase(Base_General)
    For BucleGrande = CMin To CMax
        'If BucleGrande = 1 Then
            Set Registro = Base.OpenRecordset("select * from cabecera" & Condicion & " order by cdate(fecha)")
        'Else
        '    Set Registro = Base.OpenRecordset("select * from cabeceraSC10" & Condicion)
        'End If
        'If BucleGrande = 1 Then
        '    Set RegLin = Base.OpenRecordset _
        '    ("select * from tickets")
        'Else
        '    Set RegLin = Base.OpenRecordset _
        '    ("select * from ticketsSC10")
        'End If
        With Registro
            If Not .EOF Then
                
                .MoveFirst
                .MoveLast
                Barra.Max = .Recordcount + 1
                Barra.Value = 0
                'c2f/caspiunza Archivo = FreeFile
                Archivo = -1
                ''''''''''''''''''''''''''''''''''
                If Right(Elfichero, 10) <> "bbtttt.dat" Then
                    If Not Continua Then
                        On Error Resume Next
                        'c2f/caspiunza
                        Archivo = FreeFile()
                        ''''''''''''''''''''
                        Open Exp_CambiaPath(Elfichero) For Output As #Archivo
                        If Err.Number <> 0 Then
                            MsgBox CargaCadena(1380) & Exp_CambiaPath(Elfichero), vbCritical
                            Exit Sub
                        End If
                        On Error GoTo 0
                    Else
                        'c2f/caspiunza On Error GoTo fincont
                        'c2f/caspiunza Open Exp_CambiaPath(Elfichero) For Append Lock Read Write As #Archivo
                        'c2f/caspiunza Open Exp_CambiaPath(Elfichero) For Append Shared As #Archivo
                        On Error Resume Next
                        
                        sSs = Exp_CambiaPath(Elfichero)
                        If Dir(sSs) = "" Then
                            Archivo = FreeFile()
                            Open Elfichero For Output As #Archivo
                            Close #Archivo
                            Archivo = FreeFile()
                        Else
                            If Archivo = -1 Then Archivo = FreeFile()
                        End If
                        Open Elfichero For Append As #Archivo
                        
                        'c2f/caspiunza On Error GoTo 0
             
                    End If
                End If
                .MoveFirst
                Do Until .EOF
                    
                    lSeHaCreado = False
                    
                    '**********************
                    ' previene acaparar el procesador
                    CuentaEvents = CuentaEvents + 1
                    If CuentaEvents >= 20 Then
                        Do_Events
                        CuentaEvents = 0
                    End If
                    '**********************
                    SabLineas = 0
                    Barra.Value = Barra.Value + 1
                    Buffer = ""
                    Contbuffer = 1
                    fragmentos = Fcabs
                    ReDim Fragmento(Fcabs + 1)
                    For BFrag = 1 To fragmentos
                        Fragmento(BFrag).Adicional = FCab(BFrag).Adicional
                        Fragmento(BFrag).fin = FCab(BFrag).fin
                        Fragmento(BFrag).id = FCab(BFrag).id
                        Fragmento(BFrag).inicio = FCab(BFrag).inicio
                    Next BFrag
                    
                    If Right(Elfichero, 10) = "bbtttt.dat" Then
                        On Error GoTo fincont
                        Open Exp_CambiaPath(Elfichero, !codcli, !vended, !seccion, !nume, !Fecha, !hora, !Balanza) For Output As #Archivo
                        lSeHaCreado = True
                        On Error GoTo 0
                    End If
                    
                    If Right(Elfichero, 10) = "bbtttt.dat" Then
                        ExportaunaCabecera Buffer, Contbuffer, Registro, BucleGrande, Continua, True
                    Else
                        ExportaunaCabecera Buffer, Contbuffer, Registro, BucleGrande, Continua, False
                    End If
                    
                    BufCab = Buffer
                    Print #Archivo, Buffer
                    CadenaEnlace = "date=" & Chr(34) & !Fecha & Chr(34) & _
                    " and val(nume)=" & Val(!nume) & " and numbal=" & Val(!Balanza) _
                    & " and val(secc)=" & Val(!seccion) & " and val(vendor)=" _
                    & Val(!vended) & " and val(stpetic)=" & Val(!STPETIC) & " and hour=" & _
                    Chr(34) & !hora & Chr(34)
                    
                    '********************************
                    ' exporta las líneas correspondientes
                    '********************************
                    'If Not Sabeco Then
                        'If BucleGrande = 1 Then
                        Set RegLin = Base.OpenRecordset _
                            ("select * from tickets where " & CadenaEnlace & Cond2 & " order by val(numlin)")
                        'Else
                        '    Set RegLin = Base.OpenRecordset _
                        '    ("select * from ticketsSC10 where " & CadenaEnlace & Cond2 & " order by val(numlin)")
                        'End If
                    'Else
                    '    If BucleGrande = 1 Then
                    '    Set RegLin = Base.OpenRecordset _
                    '        ("select * from tickets where " & CadenaEnlace & Cond2 & " order by val(lincan),val(numlin)")
                    '    Else
                    '        Set RegLin = Base.OpenRecordset _
                    '        ("select * from ticketsSC10 where " & CadenaEnlace & Cond2 & " order by val(lincan),val(numlin)")
                    '    End If
                    'End If
                    fragmentos = FLins
                    ReDim Fragmento(FLins + 1)
                    For BFrag = 1 To fragmentos
                        Fragmento(BFrag).Adicional = FLin(BFrag).Adicional
                        Fragmento(BFrag).fin = FLin(BFrag).fin
                        Fragmento(BFrag).id = FLin(BFrag).id
                        Fragmento(BFrag).inicio = FLin(BFrag).inicio
                    Next BFrag
                    If Not RegLin.EOF Then
                        '***********************
                        ' SABECO :
                        ' pondrá los números
                        ' de línea canceladas
                        ' a continuación de los
                        ' no cancelados"
                        '***********************
                        'If Sabeco And SabLineas = 0 Then
                        '    SabLineas = 1
                        '    RegLin.MoveFirst
                        '    Do Until RegLin.EOF
                        '        If RegLin!lincan = 0 Then SabLineas = SabLineas + 1
                        '        RegLin.Movenext
                        '    Loop
                        'End If
                        '************************
                        With RegLin
                            .MoveFirst
                            Do Until .EOF
                                Contbuffer = 1
                                Buffer = ""
                                BufferBis = ""
                                If Right(Elfichero, 10) = "bbtttt.dat" Then
                                    Call ExportaunaLinea(Buffer, BufferBis, Contbuffer, RegLin, BucleGrande, BufCab, Continua, True)
                                Else
                                    Call ExportaunaLinea(Buffer, BufferBis, Contbuffer, RegLin, BucleGrande, BufCab, Continua, False)
                                End If
                                Print #Archivo, Buffer
                                If Trim(BufferBis) <> "" Then
                                    Print #Archivo, BufferBis
                                End If
                                .Movenext
                            Loop
                        End With
                        If Right(Elfichero, 10) = "bbtttt.dat" And lSeHaCreado Then
                           lSeHaCreado = False
                           Close #Archivo
                        End If
                    End If
                    .Movenext
                Loop
                
                'c2f/caspiunza
                If Archivo <> -1 Then
                    Close #Archivo
                End If
                '''''''''''''''
                
        End If
    End With
    Next BucleGrande
    'c2f/caspiunza Close #Archivo
    'If MostrarMensajes Then MsgBox CargaCadena( 849)
    Unload Me
    Exit Sub
fincont:
    Exit Sub
errorlectura:
    MsgBox CargaCadena(1032), vbCritical + vbOKOnly
End Sub

Private Sub ExportaunaCabecera(Buffer As String, ByRef Contbuffer As Long, Registro As dao.Recordset, _
          ByRef BucleGrande As Long, Optional Continua As Boolean, Optional lEspecial As Boolean)
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim MiValor As Double
    Dim sM As String
    
    'On Error Resume Next
    'If Not lEspecial Then
    '   Dim lEspecial As Boolean
    '   lEspecial = False
    'End If
    'On Error GoTo 0
    If IsNull(lEspecial) Then lEspecial = False
    
    With Registro
    For bucle = 1 To fragmentos
        Do While Contbuffer < Fragmento(bucle).inicio
            Contbuffer = Contbuffer + 1
            'Buffer = Buffer & " "
        Loop
        laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
        Select Case Val(Fragmento(bucle).id)
            Case 0
                Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 1
                'If TipoExport <> 3 Then
                    Buffer = Buffer & Formatea(Val(!nume), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'Else
                '    If BucleGrande = 1 Then
                '        Buffer = Buffer & GA
                '    Else
                '        Buffer = Buffer & SC10
                '    End If
                '    Buffer = Buffer & Formatea(Val(!nume), laLongitud - 1, True)
                '    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 2
                Buffer = Buffer & Formatea(Val(!seccion), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 3
                Buffer = Buffer & Formatea(!Balanza, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 4
                Buffer = Buffer & Formatea(Val(!vended), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 5
                'If Sabeco Then
                '    Buffer = Buffer & Formatea(Val(!numlin) + Val(!lincan), laLongitud, True)
                '    Contbuffer = Contbuffer + laLongitud
                'Else
                    Buffer = Buffer & Formatea(Val(!numlin), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 6
                Buffer = Buffer & Formatea(Val(!lincan), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 7
                If Not lEspecial Then
                    If UsaEuro Then
                        MiValor = !Importe * 100
                    Else
                        MiValor = !Importe * (10 ^ decimales)
                    End If
                End If
                If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                If lEspecial Then
                    sM = CStr(!Importe)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    'sM = Mid("00000000000000000000", 1, LaLongitud - Len(sM)) & sM
                    Buffer = Buffer & sM 'Mid("00000000000000000000", 1, LaLongitud - 3) & ",00")
                Else
                    Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 8
                Buffer = Buffer & Formatea(FormHora(!hora), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 9
                'If Not Sabeco Then
                    Buffer = Buffer & Formatea(Val(!STPETIC), laLongitud, True)
                'Else
                '    If (Val(!STPETIC) And 64) Or (Val(!STPETIC) And 128) Then
                '        Buffer = Buffer & Formatea(1, laLongitud, True)
                '    Else
                '        Buffer = Buffer & Formatea(0, laLongitud, True)
                '    End If
                'End If
                Contbuffer = Contbuffer + laLongitud
            Case 10
                Buffer = Buffer & Formatea(!codcli, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 11
                Buffer = Buffer & Formatea(FormFecha(!Fecha), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 12
                Buffer = Buffer & Formatea(Val(!TYPTIC), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 13
                If Trim(!codbar) <> "" Then
                    Buffer = Buffer & Formatea(Mid(!codbar, Val(left(Fragmento(bucle).Adicional, 2)), Val(Right(Fragmento(bucle).Adicional, 2))), laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Space(13), laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 14
                If !Importe >= 0 Then
                    Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 15
                Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 16 'c2f invicta forma de pago
                If Not IsNull(!tipoPago) Then
                    Buffer = Buffer & Formatea(Val(!tipoPago), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 17 'c2f invicta importe bruto
                If UsaEuro Then
                    MiValor = !importe_bruto * 100
                Else
                    MiValor = !importe_bruto * (10 ^ decimales)
                End If
                If MiValor < 0 And Exp_Absoluto Then MiValor = MiValor * (-1)
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 18 'c2f lote
                Buffer = Buffer & Formatea(!numlote, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 19 'c2f factura
                On Error Resume Next
                Buffer = Buffer & Formatea(!factura, laLongitud, False)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Buffer = Buffer & Formatea(" ", laLongitud, False)
                End If
                On Error GoTo 0
                Contbuffer = Contbuffer + laLongitud
            Case 20 '% 0
                MiValor = !por0 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 21 '% 1
                MiValor = !por1 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 22 '% 2
                MiValor = !por2 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 23 '% 3
                MiValor = !por3 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 24 '% 4
                MiValor = !por4 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
        
            Case 25 'base0
                MiValor = !base0 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 26 '1
                MiValor = !base1 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 27 '2
                MiValor = !base2 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 28 '3
                MiValor = !base3 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 29 '4
                MiValor = !base4 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
        
            Case 30 'imp0
                MiValor = !imp0 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 31 '1
                MiValor = !imp1 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 32 '2
                MiValor = !imp2 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 33 '3
                MiValor = !imp3 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 34 '4
                MiValor = !imp4 * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
                
            Case 35
                MiValor = !descuento * 100
                sM = Formatea(MiValor, laLongitud, True)
                If IsNull(sM) Or Trim(sM) = "" Then
                    sM = Formatea(0, laLongitud, True)
                End If
                Buffer = Buffer & sM 'Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 36 'T. Units
                Buffer = Buffer & Formatea(!TotalUnits, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 37
                If IsNull(!TotalWeight) Then
                    MiValor = 0
                Else
                    MiValor = !TotalWeight * 1000
                End If
                sM = Formatea(MiValor, laLongitud, True)
                If IsNull(sM) Or Trim(sM) = "" Then
                    sM = Formatea(0, laLongitud, True)
                End If
                Buffer = Buffer & sM 'Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
                
        End Select
    Next bucle
    'If Continua Then
        Edit_Record Registro
        !reserv3 = "SI"
        .Update
    'End If
    End With
End Sub
'hydratouch actual 1.0.0
Private Sub ExportaunaLinea(Buffer As String, BufferBis As String, ByRef Contbuffer As Long, Registro As dao.Recordset, ByRef BucleGrande As Long, BufCab As String, Continua As Boolean, lEspecial As Boolean)
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim MiLen As Integer
    Dim sM As String
    Dim MiValor
    Dim ttt As String
    'peso/unidad, precio e importe cambio signoo para BufferBis --> 2ºlínea cancelados
    
    'On Error Resume Next
    'If Not lEspecial Then
    '   Dim lEspecial As Boolean
    '   lEspecial = False
    'End If
    'On Error GoTo 0
    If IsNull(lEspecial) Then lEspecial = False
    
    'c2f decepal
    If Dir(App.Path & "\decepal.txt") <> "" Then 'Taquion And TaquionFichero = "" And sunTicket = "" Then
        Continua = True
    End If
    
    With Registro
    
    For bucle = 1 To fragmentos
        Do While Contbuffer < Fragmento(bucle).inicio
            Contbuffer = Contbuffer + 1
            'Buffer = Buffer & " "
            'BufferBis = BufferBis & " "
        Loop
        laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
        Select Case Val(Fragmento(bucle).id)
            Case 0
                Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                BufferBis = BufferBis & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 1
                Buffer = Buffer & Formatea(FormFecha(!Date), laLongitud, False)
                BufferBis = BufferBis & Formatea(FormFecha(!Date), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 2
                Buffer = Buffer & Formatea(FormHora(!Hour), laLongitud, False)
                BufferBis = BufferBis & Formatea(FormHora(!Hour), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 3
                Buffer = Buffer & Formatea(Val(!Vendor), laLongitud, True)
                BufferBis = BufferBis & Formatea(Val(!Vendor), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 4
                'If TipoExport <> 3 Then
                    Buffer = Buffer & Formatea(Val(!nume), laLongitud, True)
                    BufferBis = BufferBis & Formatea(Val(!nume), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'Else
                '    If BucleGrande = 1 Then
                '        Buffer = Buffer & GA
                '        BufferBis = BufferBis & GA
                '    Else
                '        Buffer = Buffer & SC10
                '        BufferBis = BufferBis & SC10
                '    End If
                '    Buffer = Buffer & Formatea(Val(!nume), laLongitud - 1, True)
                '    BufferBis = BufferBis & Formatea(Val(!nume), laLongitud - 1, True)
                '    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 5
                'If Sabeco And !lincan = 1 Then
                '    Buffer = Buffer & Formatea(SabLineas, laLongitud, True)
                '    Contbuffer = Contbuffer + laLongitud
                '    SabLineas = SabLineas + 1
                'Else
                    Buffer = Buffer & Formatea(Val(!numlin), laLongitud, True)
                    BufferBis = BufferBis & Formatea(Val(!numlin), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 6
                Buffer = Buffer & Formatea(Val(!secc), laLongitud, True)
                BufferBis = BufferBis & Formatea(Val(!secc), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 7
                If Dir(App.Path & "\decepal.txt") <> "" Then
                    Buffer = Buffer & FormateaR(!Code, laLongitud, False) '  Formatea(Val(!Code), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(Val(!Code), laLongitud, True)
                End If
                BufferBis = BufferBis & Formatea(Val(!Code), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 8
                If Not lEspecial Then
                If UsaEuro Then
                    Buffer = Buffer & Formatea(!Price * 100, laLongitud, True)
                    BufferBis = BufferBis & Formatea((!Price * 100) * (-1), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(!Price * (10 ^ decimales), laLongitud, True)
                    BufferBis = BufferBis & Formatea((!Price * (10 ^ decimales)) * (-1), laLongitud, True)
                End If
                Else
                    sM = CStr(!Price)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 9
                If Not lEspecial Then
                If UsaEuro Then
                    Buffer = Buffer & Formatea(!amount * 100, laLongitud, True)
                    BufferBis = BufferBis & Formatea((!amount * 100) * (-1), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(!amount * (10 ^ decimales), laLongitud, True)
                    BufferBis = BufferBis & Formatea((!amount * (10 ^ decimales)) * (-1), laLongitud, True)
                End If
                Else
                    sM = CStr(!amount)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                    BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 10
                If Not lEspecial Then
                Buffer = Buffer & Formatea(!Weight * 1000, laLongitud, True)
                BufferBis = BufferBis & Formatea((!Weight * 1000) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!Weight)
                    If IsNull(sM) Or sM = "0" Then sM = "0,000"
                    sM = formatic(sM, 3, laLongitud)
                    Buffer = Buffer & sM
                    BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 11
                Buffer = Buffer & Formatea(!units, laLongitud, True)
                BufferBis = BufferBis & Formatea((!units) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Contbuffer = Contbuffer + laLongitud
            Case 12
                If !units = 0 Then
                    If Not lEspecial Then
                        Buffer = Buffer & Formatea(!Weight * 1000, laLongitud, True)
                        BufferBis = BufferBis & Formatea((!Weight * 1000) * (-1), laLongitud, True)
                        'BufferBis = Buffer
                    Else
                        sM = CStr(!Weight)
                        If IsNull(sM) Or sM = "0" Then sM = "0,000"
                        sM = formatic(sM, 3, laLongitud)
                        Buffer = Buffer & sM
                        BufferBis = BufferBis & sM
                    End If
                Else
                    'If Not Sabeco Then
                        Buffer = Buffer & Formatea(!units, laLongitud, True)
                        BufferBis = BufferBis & Formatea((!units) * (-1), laLongitud, True)
                        'BufferBis = Buffer
                    'Else
                    '    Buffer = Buffer & Formatea(!units * 1000, laLongitud, True)
                    'End If
                End If
                'BufferBis = Buffer
                Contbuffer = Contbuffer + laLongitud
            Case 13
                Buffer = Buffer & Formatea(!TYPTIC, laLongitud, True)
                BufferBis = BufferBis & Formatea(!TYPTIC, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 14
                Buffer = Buffer & Formatea(Val(!STPETIC), laLongitud, True)
                BufferBis = BufferBis & Formatea(Val(!STPETIC), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 15
                If !amount >= 0 Then
                    Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                End If
                BufferBis = BufferBis & Formatea(Exp_Mas, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 16
                MiLen = 1 + Val(Right(Fragmento(bucle).Adicional, 2)) - Val(left(Fragmento(bucle).Adicional, 2))
                If BufCab = "" Then 'forma de pago LLEGA NULO
                    If IsNull(!tipoPago) Then
                        Buffer = Buffer & Formatea(0, laLongitud, True)
                        BufferBis = BufferBis & Formatea(0, laLongitud, True)
                    Else
                        Buffer = Buffer & Formatea((!tipoPago), laLongitud, True)
                        BufferBis = BufferBis & Formatea((!tipoPago), laLongitud, True)
                    End If
                Else
                    If Val(Right(Fragmento(bucle).Adicional, 2)) <= Len(BufCab) Then
                        Buffer = Buffer & Formatea(Mid(BufCab, Val(left(Fragmento(bucle).Adicional, 2)), MiLen), laLongitud, False)
                        BufferBis = BufferBis & Formatea(Mid(BufCab, Val(left(Fragmento(bucle).Adicional, 2)), MiLen), laLongitud, False)
                    Else
                        Buffer = Buffer & Formatea(Space(MiLen), laLongitud, False)
                        BufferBis = BufferBis & Formatea(Space(MiLen), laLongitud, False)
                    End If
                End If
                Contbuffer = Contbuffer & laLongitud
            Case 17
                If Not IsNull(!Item) Then
                    Buffer = Buffer & Formatea(!Item, laLongitud, False)
                    BufferBis = BufferBis & Formatea(!Item, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Space(25), laLongitud, False)
                    BufferBis = BufferBis & Formatea(Space(25), laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 18
                If !units <> 0 Then
                    Buffer = Buffer & Formatea("1", laLongitud, False) '¡¡ era 2 C2F
                    BufferBis = BufferBis & Formatea("1", laLongitud, False) '¡¡ era 2 C2F
                Else
                    If !Weight <> 0 Then
                        Buffer = Buffer & Formatea("0", laLongitud, False) '¡¡ era 1 C2F
                        BufferBis = BufferBis & Formatea("0", laLongitud, False) '¡¡ era 1 C2F
                    Else
                        Buffer = Buffer & Formatea("1", laLongitud, False) '¡¡ era 0 C2F
                        BufferBis = BufferBis & Formatea("1", laLongitud, False) '¡¡ era 0 C2F
                    End If
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 19
                Buffer = Buffer & Formatea(!lincan, laLongitud, False)
                BufferBis = BufferBis & Formatea(!lincan, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 20
                If Not IsNull(!Plu) Then
                    Buffer = Buffer & Formatea(Val(!Plu), laLongitud, True)
                    BufferBis = BufferBis & Formatea(Val(!Plu), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                    BufferBis = BufferBis & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 21
                If Not IsNull(!numbal) Then
                    Buffer = Buffer & Formatea(!numbal, laLongitud, True)
                    BufferBis = BufferBis & Formatea(!numbal, laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                    BufferBis = BufferBis & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 22
                Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                BufferBis = BufferBis & Formatea(TiendaActual, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 23 'c2f invicta --> Departamento (familia)
                If Not IsNull(!DPT) Then
                    Buffer = Buffer & Formatea(Val(!DPT), laLongitud, True)
                    BufferBis = BufferBis & Formatea(Val(!DPT), laLongitud, True)
                Else
                    BufferBis = BufferBis & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 24 'c2f invicta importe bruto LLEGA NULO
                'If UsaEuro Then
                    If IsNull(!importe_bruto) Then
                        Buffer = Buffer & Formatea(0, laLongitud, True)
                        BufferBis = BufferBis & Formatea(0, laLongitud, True)
                    Else
                        Buffer = Buffer & Formatea(!importe_bruto * 100, laLongitud, True)
                        BufferBis = BufferBis & Formatea(!importe_bruto * 100, laLongitud, True)
                    End If
                'Else
                '    Buffer = Buffer & Formatea(!importe_bruto * (10 ^ decimales), laLongitud, True)
                'End If
                Contbuffer = Contbuffer + laLongitud
            Case 25 'c2f lote 'LLEGA NULO
                If IsNull(!numlote) Then
                    Buffer = Buffer & Space(laLongitud)
                    BufferBis = BufferBis & Space(laLongitud)
                Else
                    Buffer = Buffer & Formatea(!numlote, laLongitud, False)
                    BufferBis = BufferBis & Formatea(!numlote, laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 26 'c2f cliente LLEGA NULO
                If IsNull(!clientet) Then
                    Buffer = Buffer & Space(laLongitud) 'Formatea(0, laLongitud, True)
                    BufferBis = BufferBis & Space(laLongitud)
                Else
                    If IsNumeric(!clientet) Then
                        ttt = Format(Val(!clientet), Mid("000000000000000", 1, laLongitud))
                    Else
                        ttt = !clientet
                    End If
                    Buffer = Buffer & Formatea(ttt, laLongitud, False) 'True)
                    BufferBis = BufferBis & Formatea(ttt, laLongitud, False) 'True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 27
                If Not lEspecial Then
                    Buffer = Buffer & Formatea(!porcentaje * 100, laLongitud, True)
                    BufferBis = BufferBis & Formatea((!porcentaje * 100) * (-1), laLongitud, True)
                'BufferBis = Buffer
                Else
                    sM = CStr(!porcentaje)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic(sM, 3, laLongitud)
                    Buffer = Buffer & sM
                    BufferBis = BufferBis & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 28 'c2f factura Simpli.
                'Buffer = Buffer & Formatea(!factura, laLongitud, False)
                'Contbuffer = Contbuffer + laLongitud
                On Error Resume Next
                Buffer = Buffer & Formatea(!factura, laLongitud, False)
                BufferBis = BufferBis & Formatea(!factura, laLongitud, False)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Buffer = Buffer & Formatea(" ", laLongitud, False)
                    BufferBis = BufferBis & Formatea(" ", laLongitud, False)
                End If
                On Error GoTo 0
                Contbuffer = Contbuffer + laLongitud
            Case 29
                MiValor = !descuento * 100
                sM = Formatea(MiValor, laLongitud, True)
                If IsNull(sM) Or Trim(sM) = "" Then
                    sM = Formatea(0, laLongitud, True)
                End If
                Buffer = Buffer & sM 'Formatea(MiValor, laLongitud, True)
                BufferBis = BufferBis & "-" & Mid(sM, 2)
                Contbuffer = Contbuffer + laLongitud
            Case 30
                Buffer = Buffer & Formatea(!IVA, laLongitud, False)
                BufferBis = BufferBis & Formatea(!IVA, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            
        End Select
    Next bucle
    
    If !lincan <> "1" Then
        BufferBis = ""
    'Else
    '    BufferBis = ""
    End If
    
    If Continua Then
        Edit_Record Registro
        !reserv3 = "SI"
        .Update
    End If
    
    End With
    
End Sub

Private Sub LeeFragmentos(Elfichero As String)
    Dim bucle As Long
    Dim Bucle2 As Long
    Dim Archivo As Integer
    Dim Buffer As String
    Dim Buf2 As String
    Dim Buf3 As String
    Dim FragmentoAux As TipoFragmento
    Dim C1 As String
    Dim C2 As String
    Dim C3 As String
    Dim C4 As String
    Dim Contador As Long
    Buf2 = ""
    For bucle = Len(Elfichero) To 1 Step -1
        If Mid(Elfichero, bucle, 1) <> "\" Then
            Buf2 = Buf2 & Mid(Elfichero, bucle, 1)
        Else
            Exit For
        End If
    Next bucle
    Buf3 = ""
    For bucle = Len(Buf2) To 1 Step -1
        Buf3 = Buf3 & Mid(Buf2, bucle, 1)
    Next bucle
    Archivo = FreeFile
    Open Miruta & "\etc\" & Buf3 & ".fet" For Input As Archivo
    Line Input #Archivo, Buffer
    If Val(Buffer) = 4 Or Val(Buffer) = 5 Then
        Input #Archivo, C1, C2, C3
        TipoExport = Val(C1)
        GA = C2
        SC10 = C3
    End If
    Contador = 0
    Do Until EOF(Archivo)
        Contador = Contador + 1
        ReDim Preserve Fragmento(Contador + 1)
        Input #Archivo, C1, C2, C3, C4
        Fragmento(Contador).inicio = Val(C1)
        Fragmento(Contador).fin = Val(C2)
        Fragmento(Contador).id = Val(left(C3, 3))
        Fragmento(Contador).Adicional = C4
    Loop
    Close #Archivo
    fragmentos = Contador
    ' ordena
    For bucle = 1 To fragmentos - 1
        For Bucle2 = bucle + 1 To fragmentos
            If Fragmento(bucle).inicio > Fragmento(Bucle2).inicio Then
                FragmentoAux = Fragmento(bucle)
                Fragmento(bucle) = Fragmento(Bucle2)
                Fragmento(Bucle2) = FragmentoAux
            End If
        Next Bucle2
    Next bucle
  
End Sub
Private Function Formatea(cadena, Longitud As Integer, EsNumerico As Boolean) As String
    Dim Buffer As String
    Dim BufferFormato As String
    Dim bucle As Long
    If EsNumerico Then
        If IsNull(cadena) Then cadena = "0"
        If Val(cadena) < 0 Then
            For bucle = 1 To Longitud - 1
                BufferFormato = BufferFormato & "0"
            Next bucle
        Else
            For bucle = 1 To Longitud
                BufferFormato = BufferFormato & "0"
            Next bucle
        End If
        Buffer = Format(cadena, BufferFormato)
    Else
        Buffer = cadena
        If Len(Buffer) < Longitud Then
            Buffer = Buffer & Space(Longitud - Len(Buffer))
        End If
        Buffer = left(Buffer, Longitud)
    End If
    If Exp_Absoluto Then
        If EsNumerico Then
            If Val(Buffer) < 0 Then
                Buffer = Right(Buffer, Len(Buffer) - 1)
                Buffer = "0" & Buffer
            End If
        End If
    End If
    If Len(Buffer) > Longitud Then
        Buffer = Right(Buffer, Longitud)
    End If
    Formatea = Buffer
End Function
Private Function FormateaR(cadena, Longitud As Integer, EsNumerico As Boolean) As String
    Dim Buffer As String
    Dim BufferFormato As String
    Dim bucle As Long
    If EsNumerico Then
        For bucle = 1 To Longitud
            BufferFormato = BufferFormato & "0"
        Next bucle
        Buffer = Format(cadena, BufferFormato)
    Else
        Buffer = cadena
        If Len(Buffer) < Longitud Then
            Buffer = Space(Longitud - Len(Buffer)) & Buffer
        End If
        Buffer = left(Buffer, Longitud)
    End If
    If Exp_Absoluto Then
        If EsNumerico Then
            If Val(Buffer) < 0 Then
                Buffer = Right(Buffer, Len(Buffer) - 1)
                Buffer = "0" & Buffer
            End If
        End If
    End If
    If Len(Buffer) > Longitud Then
        Buffer = Right(Buffer, Longitud)
    End If
    FormateaR = Buffer
End Function

Private Function FormFecha(MiFecha As String) As String
    Dim Retorno As String
    Dim Dia As String
    Dim Mes As String
    Dim Anyo As String
    Dia = left(MiFecha, 2)
    Mes = Mid(MiFecha, 4, 2)
    Anyo = Right(MiFecha, 2)
    '**************
    ' mifecha --> ddmmyy
    '**************
    Select Case Exp_Fecha
        Case 0
            Retorno = Dia & "/" & Mes & "/" & Anyo
        Case 1
            Retorno = Mes & "/" & Dia & "/" & Anyo
        Case 2
            Retorno = Dia & "/" & Mes & "/" & "20" & Anyo
        Case 3
            Retorno = Mes & "/" & Dia & "/" & "20" & Anyo
        Case 4
            Retorno = Dia & Mes & Anyo
        Case 5
            Retorno = Mes & Dia & Anyo
        Case 6
            Retorno = Dia & Mes & "20" & Anyo
        Case 7
            Retorno = Mes & Dia & "20" & Anyo
        Case 8
            Retorno = Dia & "-" & Mes & "-" & Anyo
    End Select
    FormFecha = Retorno
End Function
Private Function FormHora(MiHora As String) As String
    
    Dim Mhora As String
    Dim MMinuto As String
    Dim Retorno As String
    '****************
    ' mihora --> hh:mm
    '****************
    Mhora = left(MiHora, 2)
    MMinuto = Mid(MiHora, 4, 2)
    Select Case Exp_hora
        Case 0
            Retorno = Mhora & ":" & MMinuto
        Case 1
            Retorno = Mhora & MMinuto
        Case 2
            Retorno = Mhora & "-" & MMinuto
        Case 3
            Retorno = Mhora
    End Select
    FormHora = Retorno
End Function
Private Function Exp_CambiaPath(MiFichero As String, Optional codcli As Variant, Optional vended As Variant, Optional _
                                seccion As Variant, Optional nume As Variant, Optional Fecha As Variant, Optional _
                                hora As Variant, Optional numbal As Variant) As String
    Dim Buffer1 As String
    Dim bucle As Integer
    Dim Contador As Integer
    Dim NombreF As Tipo_NombreTiquet
    'ver ficheros belros
    'fin_dbal.tmp
    'Tick_art.tmp
    'tick_tot.tmp
    If MiFichero = "fin_dbal.tmp" Or MiFichero = "Tick_art.tmp" Or MiFichero = "tick_tot.tmp" Then
        'Buffer1 = "." & left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
        Buffer1 = "." & Format(Day(Date), "000")
        MiFichero = Replace(MiFichero, ".tmp", Buffer1)
    End If
    
    If Right(MiFichero, 10) = "bbtttt.dat" Then
        NombreF.cliente = codcli
        NombreF.vendedor = Val(vended)
        NombreF.seccion = Val(seccion)
        NombreF.tiquet = Val(nume)
        NombreF.Balanza = Val(numbal)
        Fecha = Format(Fecha, "ddmmyyyy")
        hora = Format(hora, "hhmm")
        Buffer1 = ObtenNombre(NombreF)
        Buffer1 = Buffer1 & "." & left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
                    
    Else
        Buffer1 = MiFichero
    End If
    
    Contador = 0
    For bucle = Len(Buffer1) To 1 Step -1
        If Mid(Buffer1, bucle, 1) <> "\" Then
            Contador = Contador + 1
        Else
            Exit For
        End If
    Next bucle
    Exp_CambiaPath = Exp_Path & "\" & MiTienda & Right(Buffer1, Contador)
            
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub TxtFecha_LostFocus()
    Dim Mes As Integer
    Dim Anyo As Integer
    Dim MIERROR As Boolean
    
    MIERROR = False
    If Len(txtfecha.TexT) <> 8 Then
        MIERROR = True
    Else
        Anyo = Mid(txtfecha.TexT, 7, 2)
        If Anyo > 70 Then
            Anyo = Anyo + 2000
        Else
            Anyo = Anyo + 1900
        End If
        Mes = Mid(txtfecha.TexT, 4, 2)
        If Mes < 1 Or Mes > 12 Then
            MIERROR = True
        Else
            If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
                If Val(Mid(txtfecha.TexT, 1, 2)) > 31 Then MIERROR = True
            Else
                If Mes = 2 Then
                    If Anyo Mod 4 = 0 Then
                        If Val(Mid(txtfecha.TexT, 1, 2)) > 29 Then MIERROR = True
                    Else
                        If Val(Mid(txtfecha.TexT, 1, 2)) > 28 Then MIERROR = True
                    End If
                Else
                    If Val(Mid(txtfecha.TexT, 1, 2)) > 30 Then MIERROR = True
                End If
            End If
        End If
    End If
    If MIERROR = True Then
        MsgBox CargaCadena(506), vbCritical
        txtfecha.SetFocus
        Exit Sub
    End If
        
End Sub

Public Function formatic(sM As String, nDec As Integer, laLo As Integer) As String
   Dim nL As Integer
   Dim nB As Integer
   Dim nD
   nL = Len(sM)
   nD = 0
   For nB = 1 To nL
          'MsgBox Format(Asc(Mid(sM, nB, 1)), "000")

       If Mid(sM, nB, 1) = "," Then
       
          nD = nB
          nB = nL + 1
       End If
   
   Next nB
   If nD <> 0 Then
       If (nL - nD) < nDec Then
          
          sM = sM & Mid("0000000000", 1, nDec - (nL - nD))
           
       End If
   
   Else
       sM = sM & "," & Mid("0000000000", 1, nDec)
   End If
   sM = Mid("00000000000000000000", 1, laLo - Len(sM)) & sM
   formatic = sM
End Function
'se ha traido de taquion_lnk
'*********************************
' Euroscale=true // SC10=false
'*********************************
Public Sub Taquion_EnviaTiquetsPendientes(Euroscale_o_SC10 As Boolean)
Dim Base As dao.Database
Dim Registro As dao.Recordset
Dim Reg2 As dao.Recordset
'**************************
' Cabecera
Dim B1 As String
Dim B2 As String
Dim Codigo_mos As String
Dim Codigo_bal As String
Dim Codigo_ven As String
Dim Cod_Bar As String
Dim Lineas As String
Dim precio As String
Dim Tipo As String
Dim Fecha As String
Dim hora As String
Dim LineaAnulada As String
'********
' Linea
Dim Num_lin As String
Dim Cod_Plu As String
Dim Cod_Cod As String
Dim Precio_Uni As String
Dim Pes_Uni As String
Dim Total As String
Dim Signo As String
Dim Tipo_Linea As String
Dim Anulado As String
'************
Dim MiEnvio As String
Dim nintentos As Integer
Dim bucle As Integer
Dim RegIVA As dao.Recordset
Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim NombreTiquet As String
Dim BufferLineas() As String
Dim MiCuentaLineas As Integer
Dim BucleLineas As Integer
Dim HayCabecera As Boolean
Dim PrecioSinIVA As Double
Dim PrecioConIVA As Double
Dim BufferIVA As String
Dim BufferIVA2 As String
Dim MiIva(5) As Tipo_Grupo_IVA
Dim BucleIVA As Integer
Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
Dim Ret
Dim nFich
Dim siP, susR, spsS
'Dim nIg As Integer

    Fichero1 = FreeFile()
    MiTabla1 = "Cabecera"
    MiTabla2 = "tickets"
    Set Base = OpenDatabase(Base_General)
    '********************
    ' carga tipos de IVA
    '********************
    Set RegIVA = Base.OpenRecordset("select * from iva")
    For bucle = 0 To 4
        RegIVA.FindFirst "tipo=" & (bucle)
        If RegIVA.EOF Or RegIVA.NoMatch Then
            MiIva(bucle).porcentaje = 0
        Else
            MiIva(bucle).porcentaje = RegIVA!Valor
        End If
    Next bucle
    '*********************
    ' /carga tipos de IVA
    '*********************
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & " order by val(nume)"
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        Registro.MoveFirst
        If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
            nFich = FreeFile()
            Open App.Path & "\getputfileftp.dat" For Input As #nFich
            'Do While Not EOF(nFich)
                Line Input #nFich, siP
                Line Input #nFich, susR
                Line Input #nFich, spsS
            'Loop
            Close #nFich
            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
            nRetConnFTP = Ret
        End If
        ' si no está enviada la cabecera, se envía
        Do Until Registro.EOF Or CI_Salir = True
            With Registro
                If !reserv3 = "NO" Or !reserv3 = "--" Then
                    HayCabecera = True
                    Codigo_mos = Format(Val(!seccion), "00")
                    Codigo_bal = Format(Val(!Balanza), "00")
                    Codigo_ven = Format(Val(!vended), "0000")
                    If !codbar <> "" Then
                        Cod_Bar = !codbar
                    Else
                        Cod_Bar = Space(13)
                    End If
                    Lineas = Format(Val(!numlin), "000")
                    If !Importe >= 0 Then
                        precio = !Importe
                        precio = Quitar_Coma(precio, 8)
                    Else
                        precio = (-1) * !Importe
                        precio = "-" & Quitar_Coma(precio, 7)
                    End If
                    ' Tipo V -> venta red
                    ' Tipo L -> venta local
                    ' Tipo A -> venta anulado red
                    ' Tipo B -> venta anulado local
                    If (!STPETIC And 64) = 0 And (!STPETIC And 128) = 0 Then
                        If (!STPETIC And 1) = 0 Then
                            Tipo = "V"
                        Else
                            Tipo = "L"
                        End If
                    Else
                        If (!STPETIC And 1) = 0 Then
                            Tipo = "A"
                        Else
                            Tipo = "B"
                        End If
                    End If
                    NombreF.Balanza = Val(!Balanza)
                    Select Case !TYPTIC
                        Case 1
                            NombreF.Modo = 1
                        Case 2
                            NombreF.Modo = 2
                        Case 3
                            NombreF.Modo = 3
                        Case 4
                            NombreF.Modo = 4
                        Case 5
                            NombreF.Modo = 5
                        Case 6
                            NombreF.Modo = 6
                        Case 7
                            NombreF.Modo = 7
                        Case 8
                            NombreF.Modo = 8
                    End Select
                    NombreF.cliente = Val(!codcli)
                    NombreF.vendedor = Val(!vended)
                    NombreF.seccion = Val(!seccion)
                    NombreF.tiquet = Val(!nume)
                    Fecha = Format(!Fecha, "ddmmyyyy")
                    hora = Format(!hora, "hhmm")
                    ' Enviar la cabecera
                    MiEnvio = "H" & _
                    Codigo_mos & Codigo_bal & Codigo_ven & Cod_Bar & _
                    Lineas & precio & Tipo & Fecha & hora
                    ' Envia la cadena
                    ReDim BufferLineas(1)
                    BufferLineas(0) = MiEnvio
                    MiCuentaLineas = 1
                    NombreTiquet = ObtenNombre(NombreF)
                    'NombreTiquet = Codigo_bal & Format(!nume, "0000") & "." & _
                    'Left(fecha, 2) & Hex$(Val(Mid(fecha, 3, 2)))
                    NombreTiquet = NombreTiquet & "." & left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
                    Edit_Record Registro
                    .Fields("reserv3") = "--"
                    .Update
                Else
                    HayCabecera = False
                End If
            End With
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & " order by val(numlin)")
            With Reg2
                
                If Not .EOF Then
                    .MoveFirst
                    For BucleIVA = 0 To 4
                        MiIva(BucleIVA).Base = 0
                        MiIva(BucleIVA).Importe = 0
                    Next BucleIVA
                    Do Until .EOF
                        With Reg2
                            Num_lin = Format(!numlin, "000")
                            Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                            Cod_Cod = Format(!Code, "000000")
                            If !Price >= 0 Then
                                Precio_Uni = Round(!Price, 2)
                                Precio_Uni = Quitar_Coma(Precio_Uni, 8)
                            Else
                                Precio_Uni = (-1) * Round(!Price, 2)
                                Precio_Uni = Quitar_Coma(Precio_Uni, 7)
                                Precio_Uni = "-" & Precio_Uni
                            End If
                            If (!lincan = 0) Then
                                If !IVA >= 0 And !IVA < 5 Then
                                    MiIva(!IVA).Importe = MiIva(!IVA).Importe + !amount
                                    MiIva(!IVA).Base = MiIva(!IVA).Base + PRsinIVA(!amount, MiIva(!IVA).porcentaje)
                                End If
                            End If
                            If !units = 0 Then
                                If !Weight >= 0 Then
                                    Pes_Uni = Round(!Weight, 3)
                                    Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                                Else
                                    Pes_Uni = (-1) * Round(!Weight, 3)
                                    Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                                End If
                                Tipo_Linea = "P"
                            Else
                                If !units > 0 Then
                                    Pes_Uni = !units
                                    'Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                                    Pes_Uni = Format(Pes_Uni, "0000000")
                                Else
                                    Pes_Uni = (-1) * !units
                                    'Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                                    Pes_Uni = "-" & Format(Pes_Uni, "000000")
                                End If
                                Tipo_Linea = "U"
                            End If
                            If !amount < 0 Then
                                Total = (-1) * Round(!amount, 2)
                                Total = "-" & Quitar_Coma(Total, 7)
                                Signo = "-"
                            Else
                                Total = Round(!amount, 2)
                                Total = Quitar_Coma(Total, 8)
                                Signo = "+"
                            End If
                            If !lincan = 0 Then
                                LineaAnulada = "N"
                            Else
                                LineaAnulada = "S"
                            End If
                            If LineaAnulada <> "S" Then
                                MiEnvio = "L" & Codigo_mos & _
                                Cod_Bar & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Total & Signo & _
                                Tipo_Linea & LineaAnulada & Cod_Cod & Format(Val(!IVA), "0")
                                ' Envia la cadena
                                ReDim Preserve BufferLineas(MiCuentaLineas + 1)
                                BufferLineas(MiCuentaLineas) = MiEnvio
                                MiCuentaLineas = MiCuentaLineas + 1
                            End If
                            Edit_Record Reg2
                            .Fields("reserv3") = "SI"
                            .Update
                            .Movenext
                        End With
                    Loop
                 
                 Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where reserv3<>" & _
                 Chr(34) & "SI" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                 " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & " order by val(numlin)")
                 If Reg2.EOF Then
                 
                     Edit_Record Registro
                     Registro.Fields("reserv3") = "SI"
                     Registro.Update
                     
                     If HayCabecera Then
                        BufferIVA = ""
                        BufferIVA2 = ""
                        PrecioSinIVA = 0
                        PrecioConIVA = 0
                        For BucleIVA = 0 To 4
                            If MiIva(BucleIVA).Base >= 0 Then
                                B1 = MiIva(BucleIVA).Base
                                BufferIVA = BufferIVA & Quitar_Coma(B1, 8)
                            Else
                                B1 = (-1) * MiIva(BucleIVA).Base
                                BufferIVA = BufferIVA & "-" & Quitar_Coma(B1, 7)
                            End If
                            PrecioSinIVA = PrecioSinIVA + MiIva(BucleIVA).Base
                            '
                            If MiIva(BucleIVA).Importe >= 0 Then
                                B2 = MiIva(BucleIVA).Importe
                                BufferIVA2 = BufferIVA2 & Quitar_Coma(B2, 8)
                            Else
                                B2 = (-1) * MiIva(BucleIVA).Importe
                                BufferIVA2 = BufferIVA2 & "-" & Quitar_Coma(B2, 7)
                            End If
                            PrecioConIVA = PrecioConIVA + MiIva(BucleIVA).Importe
                        Next BucleIVA
                        '
                        If PrecioSinIVA > 0 Then
                            B1 = PrecioSinIVA
                            BufferLineas(0) = BufferLineas(0) & Quitar_Coma(B1, 8) & BufferIVA
                        Else
                            B1 = PrecioSinIVA * (-1)
                            BufferLineas(0) = BufferLineas(0) & "-" & Quitar_Coma(B1, 7) & BufferIVA
                        End If
                        '
                        BufferLineas(0) = BufferLineas(0) & BufferIVA2
                    End If
                    Fichero1 = FreeFile()
                    On Error GoTo 0
                    '*************************************************************************************
                    '*************************************************************************************cas.v160
                    If (Dir(App.Path & "\hiper", vbArchive) <> "") And (Not (bunTicket)) Then
                       sruta = App.Path & "\exporta"
                    Else
                       sruta = Miruta
                    End If
                    'If Dir(App.Path & "\ahorramas.cfg") <> "" Then
                       sruta = App.Path & "\exporta"
                    'End If
                     'Open MiruTa & "\tiquet.buffer" For Output As #Fichero1
                     Open sruta & "\tiquet.buffer" For Output As #Fichero1
                     For BucleLineas = 0 To MiCuentaLineas - 1
                        Print #Fichero1, BufferLineas(BucleLineas)
                     Next BucleLineas
                     Close #Fichero1
                     frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905) & " " & Cod_Bar
                     If DebugActivo Then CadenadeLog Format(Now, "hh:mm:ss") & " " & CargaCadena(905) & " " & Cod_Bar
                     'FileCopy MiruTa & "\tiquet.buffer", MiruTa & "\" & NombreTiquet
                     FileCopy sruta & "\tiquet.buffer", sruta & "\" & NombreTiquet
                    '*************************************************************************************
                    Do_Events
                    
                    If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 1 Then
                        'If Ret = 1 Then 'ha sido posible conectar
                            Ret = Enviar_Fichero_FTP(sruta & "\tiquet.buffer", NombreTiquet, 0)
                        'End If
                        Do_Events
                        Ret = Err.LastDllError
                        If Ret <> 0 Then 'error en DLL
                            Ret = Desconectar_FTP(0)
                            FTP_hOpen(0) = 0
                            FTP_HandleConect(0) = 0
                            nRetConnFTP = 0
                            If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
                                nFich = FreeFile()
                                Open App.Path & "\getputfileftp.dat" For Input As #nFich
                                'Do While Not EOF(nFich)
                                    Line Input #nFich, siP
                                    Line Input #nFich, susR
                                    Line Input #nFich, spsS
                                'Loop
                                Close #nFich
                                Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                                nRetConnFTP = Ret
                                If Ret = 1 Then
                                    Ret = Enviar_Fichero_FTP(sruta & "\tiquet.buffer", NombreTiquet, 0)
                                End If
                            End If

                            Ret = Err.LastDllError
                            If Ret <> 0 Then 'error en DLL
                                Ret = Desconectar_FTP(0)
                                FTP_hOpen(0) = 0
                                FTP_HandleConect(0) = 0
                                nRetConnFTP = 0
                                If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
                                    nFich = FreeFile()
                                    Open App.Path & "\getputfileftp.dat" For Input As #nFich
                                    'Do While Not EOF(nFich)
                                        Line Input #nFich, siP
                                        Line Input #nFich, susR
                                        Line Input #nFich, spsS
                                    'Loop
                                    Close #nFich
                                    Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                                    nRetConnFTP = Ret
                                    If Ret = 1 Then
                                        Ret = Enviar_Fichero_FTP(sruta & "\tiquet.buffer", NombreTiquet, 0)
                                    End If
                                End If
    
                            End If
                        
                        End If

                    End If
                    '*************************************************************************************
                     On Error GoTo 0
                 End If
            Else
                Edit_Record Registro
                Registro.Fields("reserv3") = "SI"
                Registro.Update
            End If
        End With
        Registro.Movenext
        Loop
    
        'If Dir(App.Path & "\getputfileftp.dat") <> "" Then
        '    If Ret = 1 Then 'ha sido posible conectar
        '        Ret = Desconectar_FTP(0)
        '    End If'
        '
        'End If
    
    End If
    '1.7.2
    'CerrarBase Base
    RegIVA.Close
    Set RegIVA = Nothing
    'Reg2.Close
    Set Reg2 = Nothing
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    ''''''''''''''''
    lCogeTiquet = True
    Exit Sub
ErrorPath:
    MsgBox CargaCadena(854) & " " & Miruta
    salir_programa
End Sub


'1.0.4 revisar taquion
Public Sub TQ_ImportaFicheros()
    Dim bucle As Integer
    Dim Archivo As Integer
    Dim MiFichero As String
    Dim EnvArt As Boolean
    Dim EnvVac As Boolean
    Dim mimax
    Dim B2 As Integer
    Dim nMiArch As Integer
    Dim sMiCad As String
    Dim sMiTmp As String
    Dim sMiFich As String
    Dim nNum As Integer
    Dim sMiFichero As String
    Dim nMy As Form
    
    Dim MyProc As Long
    Dim MyResul As Long
    Dim sZip As String
    Dim nResp As Integer
    Dim nFich As Integer
    Dim siP As String
    Dim susR As String
    Dim spsS As String
    
    'On Error Resume Next
    ''1.0.4 revisar para introducir hasta que finalice el proceso...
    'If Dir(App.Path & "\transorden.exe") <> "" Then
    '    Shell ("transorden.exe")
    'End If
    'On Error GoTo 0
    
    sMiFichero = ""
    
    '1.0.4 se modifica para importar:
    ' - Fampos.dat
    ' - Venpos.dat
    ' - Clientes.dat
    ' - Fichero de artículos definido...
    'If UsaFichasVacuno Then
    '    mimax = 2
    'Else
    '    mimax = 1
    'End If
    'If Dir(App.Path & "\clientes.dat") <> "" Then
    '    mimax = mimax + 1
    'End If
    'If Dir(App.Path & "\fampos.dat") <> "" Then
    '    mimax = mimax + 1
    'End If
    'If Dir(App.Path & "\venpos.dat") <> "" Then
    '    mimax = mimax + 1
    'End If
    ''''''''''''''''''''''''''''''''''
    '1.0.4 revisar
    'If Dir(App.Path & "\pathimp.txt") <> "" Then
    '    mimax = 1
    '    On Error Resume Next
    '    nMiArch = FreeFile()
    '    Open App.Path & "\pathimp.txt" For Input As #nMiArch
    '    Line Input #nMiArch, sMiCad
    '    sMiCad = Trim(sMiCad)
    '    Close #nMiArch
    '    sMiFichero = separaFich(sMiCad)
    '    'If Right(sMiCad, 1) = "\" Then sMiCad = Mid(sMiCad, 1, Len(sMiCad) - 1)
    '    On Error GoTo 0
    'End If
    ''''''''''''
    
    EnvArt = False
    EnvVac = False
    'FileTQ.Path = MiruTa & "\etc\"
    'For B2 = 1 To mimax
    '    If B2 = 1 Then
    '        If TaquionFichero <> "" Then
    '            FileTQ.Pattern = TaquionFichero & ".fia"
    '        Else
    '            FileTQ.Pattern = "*.fia"
    '        End If
    '        FileTQ.Refresh
    '    Else
    '        If UsaFichasVacuno And B2 = 2 Then
    '            FileTQ.Path = MiruTa
    '            FileTQ.Pattern = "fichas.dat"
    '            FileTQ.Refresh
    '        Else
    '            FileTQ.Path = MiruTa
    '            FileTQ.Pattern = "clientes.dat"
    '            FileTQ.Refresh
    '
    '        End If
    '
    '    End If
    '    FileTQ.Refresh
    '    sMiTmp = MiruTa
    '    If Dir(App.Path & "\pathimp.txt") <> "" Then
    '        nNum = 0
    '    Else
    '        nNum = FileTQ.ListCount - 1
    '    End If
    '
        
        '1.0.4
        sMiTmp = Miruta
        For bucle = 1 To 6
            If ((bucle = 1 And Dir(Miruta & "\fampos.dat") <> "") _
                Or (bucle = 2 And Dir(Miruta & "\venpos.dat") <> "") _
                Or (bucle = 3 And Dir(Miruta & "\artcon.dat") <> "") _
                Or (bucle = 4 And Dir(Miruta & "\clientes.dat") <> "") _
                Or (bucle = 5 And Dir(Miruta & "\" & TaquionFichero) <> "") _
                Or (bucle = 4 And Dir(Miruta & "\clientesx.dat") <> "") _
                Or (bucle = 6 And Dir(Miruta & "\bal3.dat") <> "")) Then
                ' comprueba que puede abrirlo
                frmEpelsa.CmdComunicaciones(3).Enabled = False
                frmEpelsa.BloquearMenues
                'If B2 = 1 Then
                '    If Supeco Then
                '        frmEpelsa.LblGA.Caption = CargaCadena(830) & " artibal"
                '    Else
                '        If Dir(App.Path & "\pathimp.txt") <> "" Then
                '            frmEpelsa.LblGA.Caption = CargaCadena(830) & " " & sMiCad & "," & CargaCadena(152)
                '        Else
                '            frmEpelsa.LblGA.Caption = CargaCadena(830) & " " & left(FileTQ.List(Bucle), Len(FileTQ.List(Bucle)) - 4) & "," & CargaCadena(152)
                '        End If
                '    End If
                'Else
                '    If UsaFichasVacuno And B2 = 2 Then
                '        frmEpelsa.LblGA.Caption = CargaCadena(830) & " " & "FICHAS.DAT" & "," & CargaCadena(152)
                '    Else
                '        frmEpelsa.LblGA.Caption = CargaCadena(830) & " " & "CLIENTES.DAT" & "," & CargaCadena(152)
                '    End If
                'End If
                'Archivo = FreeFile()
                'On Error GoTo siguiente
                'If B2 = 1 Then
                '    If Supeco Then
                '        Open MiruTa & "\artibal" For Input As #Archivo
                '    Else
                '        If Dir(App.Path & "\pathimp.txt") <> "" Then
                '            FileCopy sMiCad, MiruTa & "\" & sMiFichero
                '            Open MiruTa & "\" & sMiFichero For Input As #Archivo
                '
                '        Else
                '            Open MiruTa & "\" & left(FileTQ.List(Bucle), Len(FileTQ.List(Bucle)) - 4) For Input As #Archivo
                '        End If
                '    End If
                'Else
                '    If UsaFichasVacuno And B2 = 2 Then
                '        Open MiruTa & "\" & "fichas.dat" For Input As #Archivo
                '    Else
                '        Open MiruTa & "\" & "clientes.dat" For Input As #Archivo
                '    End If
                'End If
                'On Error GoTo 0
                'Close #Archivo
                    
                Set nMy = New frmControl
                nMy.ControlImport1.FamFijas = FamiliasFijas
                nMy.ControlImport1.SubFijas = SubSeccionesFijas
                nMy.ControlImport1.RechazarPrecioCero = RechazarPlu0
                nMy.ControlImport1.PathBase = Base_General
                nMy.ControlImport1.FamVacuno = FamiliaEspecial
                nMy.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                nMy.ControlImport1.PathFile = Miruta
                nMy.ControlImport1.PathNTQ = Miruta
                Select Case bucle
                Case 1
                sMiFichero = "fampos.dat"
                nMy.ControlImport1.Fichero = "fampos.dat"
                Case 2
                sMiFichero = "venpos.dat"
                nMy.ControlImport1.Fichero = "venpos.dat"
                Case 3
                sMiFichero = "artcon.dat"
                nMy.ControlImport1.Fichero = "artcon.dat"
                Case 4
                If Dir(Miruta & "\clientes.dat") <> "" Then
                    sMiFichero = "clientes.dat"
                    nMy.ControlImport1.Fichero = "clientes.dat"
                Else
                    sMiFichero = "clientes.dat"
                    nMy.ControlImport1.Fichero = "clientesx.dat"
                End If
                Case 5
                sMiFichero = TaquionFichero
                nMy.ControlImport1.Fichero = TaquionFichero
                Case 4
                sMiFichero = "bal3.dat"
                nMy.ControlImport1.Fichero = "bal3.dat"
                
                End Select
                nMy.ControlImport1.MuestraMensajes = False
                nMy.ControlImport1.ActivarLog = DebugActivo
                nMy.ControlImport1.MostrarFormulario = False
                nMy.ControlImport1.FactorEuro = 1
                nMy.ControlImport1.NumeroDecimales = decimales
                If UsaEuro Then
                    nMy.ControlImport1.Euros = True
                Else
                    nMy.ControlImport1.Euros = False
                End If
                TqStatus = 1
                frmEpelsa.CmdComunicaciones(4).Caption = CargaCadena(913)
                frmEpelsa.CmdComunicaciones(4).Visible = True
                nMy.ControlImport1.Importar_Background sMiFichero
                Do Until nMy.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                nMy.ControlImport1.CancelarImportacion
                Unload nMy
                Set nMy = Nothing
                Corregir_Nulos
                frmEpelsa.LblGA.Caption = CargaCadena(910)
                If Dir(App.Path & "\raro") = "" Then
                    Kill Miruta & "\" & sMiFichero
                End If
                '''resultado importación artículos
                
                MyResul = InStr(1, sMiFichero, ".")
                If MyResul <> 0 Then
                    sMiCad = Mid(sMiFichero, 1, MyResul - 1) & ".log"
                    sZip = Mid(sMiFichero, 1, MyResul - 1) & ".zip"
                Else
                    sMiCad = sMiFichero & ".log"
                    sZip = sMiFichero & ".zip"
                End If
                If Dir(App.Path & "\" & sMiCad) <> "" Then
                    If Dir(Miruta & "\" & sZip) <> "" Then
                        Kill Miruta & "\" & sZip
                    End If
                
                    MyResul = InStr(1, sMiFichero, ".")
                    MyResul = STILL_ACTIVE
                    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Miruta & "\" & "zip " & sZip & " " & sMiCad, vbMinimized))
                    Do While MyResul = STILL_ACTIVE
                        GetExitCodeProcess MyProc, MyResul
                        Do_Events
                    Loop
                    If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
                        nFich = FreeFile()
                        Open App.Path & "\getputfileftp.dat" For Input As #nFich
                        'Do While Not EOF(nFich)
                            Line Input #nFich, siP
                            Line Input #nFich, susR
                            Line Input #nFich, spsS
                        'Loop
                        Close #nFich
                        MyResul = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                        nRetConnFTP = MyResul
                    End If
                    If nRetConnFTP <> 0 Then
                        MyResul = Enviar_Fichero_FTP(App.Path & "\" & sZip, Format(Now(), "ddmmyy") & sZip, 0)
                    End If
                End If
                ''''''''''''''''''''''''''''''''''
            End If
        
            
        Next bucle
    
    If Dir(App.Path & "\asign.tst") <> "" Then
        Call AsignTec
    End If
    
    'Next B2
    
    'Set frmControl = Nothing
    
    If PausaTiquets = True Then frmEpelsa.HabilitarMenues
    frmEpelsa.CmdComunicaciones(3).Enabled = True
    'If frmControl.FileTQ.ListCount > 0 Then
    frmEpelsa.CmdComunicaciones(3).Enabled = False
    TqStatus = 2
    
    'Colomo
    'if EnvArt = True And TqDirecto Then
    If TqDirecto Then
        frmEpelsa.CmdComunicaciones(4).Caption = CargaCadena(911)  '"Detener Envío de Artículos"
        frmEpelsa.CmdComunicaciones(4).Visible = True
        'FrmExportar.AN_Articulos True
        AN_BaseMaestra True, , 7 'modificaciones artículos.
        'If (Not usaSC10 Or (usaSC10 And SC10Rs232)) Then FrmExportar.AN_TeclasRapidas True
    End If
    If EnvVac = True And TqDirecto Then
        frmEpelsa.CmdComunicaciones(4).Caption = CargaCadena(912)
        frmEpelsa.CmdComunicaciones(4).Visible = True
        'FrmExportar.AN_FichasVacuno True
        AN_BaseMaestra True, , 5
    End If
    
    
    frmEpelsa.CmdComunicaciones(3).Enabled = True
    'End If
    TqStatus = 0
    frmEpelsa.CmdComunicaciones(4).Visible = False
    
    
    Miruta = sMiTmp
    
    'c2f/caspiunza **
    If Dir(App.Path & "\orden.dat") <> "" Then
        Kill App.Path & "\orden.dat"
    End If
    If Dir(App.Path & "\ordtec.dat") <> "" Then
        Kill App.Path & "\ordtec.dat"
    End If
    
    
    Do_Events
    
    'c2f/caspiunza **
    'Unload frmControl
    'Do_Events
    'Unload frmControl
    'Do_Events
    'Set frmControl = Nothing
    'Do_Events
    ''''''''''''''
       
    Exit Sub
siguiente:
    
    Miruta = sMiTmp
    
    If Supeco Then
        If Dir(App.Path & "\art.dat") <> "" Then
            On Error Resume Next
            Kill App.Path & "\art.dat"
            On Error GoTo 0
        End If
    End If
    'c2f/caspiunza
    If Dir(App.Path & "\orden.dat") <> "" Then
        Kill App.Path & "\orden.dat"
    End If
    If Dir(App.Path & "\ordtec.dat") <> "" Then
        Kill App.Path & "\ordtec.dat"
    End If
    ''''''''''''''
    
    If Not Supeco Then frmEpelsa.LblGA.Caption = CargaCadena(854) & " " & CargaCadena(633)
End Sub

'1.7.2 las funciones que siguen estaban en envios_alto_nivel.bas
'////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////
Public Sub AN_FinDeDia(Optional ByVal lBorrado As Boolean)
Dim sDia As String
Dim sMes As String
Dim sext As String
Dim sruta As String
Dim OkBor As Long

    lCogeTiquet = False
    FrmExportar.AN_RecogerTotales (lBorrado)
    
    If lBorrado Then

        HabilitarSalir
        'If (Taquion And TaquionTiquets) Or AhorraMas Then
        '    frmEpelsa.Enabled = False
        '    frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & CargaCadena(707)
        '    FrmExportar.Taquion_EnviaTiquetsPendientes True
        '    '*************************************************************************************cas.v160->3º paso
        '    If AhorraMas And (Not (bunTicket)) Then
        '        sruta = App.Path & "\exporta"
        '        On Error Resume Next
        '        Kill sruta & "\*.*"
        '        On Error GoTo 0
        '    End If
        '    '*******cas.v160
        '    frmEpelsa.Enabled = True
        'End If
        grventas
        lCogeTiquet = False
        '**************
        ' fin dia (exporta)
        '**************
        Trata_Fin_Dia
        '*************************
        ' xmsb100, genera totales.dat
        '*************************
        If (lClR And sPathMaj <> "") Or Dir(Miruta & "\xmsb100.exe") <> "" Or Dir(Miruta & "\fedeiv8.exe") <> "" Or Dir(Miruta & "\fede386.exe") <> "" Then
            ExportaTotalesDat False
        End If
        If UsaBackup Then
            FrmResturarBackup.Realizar_Backup False
        End If
        
        'If Not Sabeco Then
            If Dir(Miruta & "\end_day.bat") <> "" Then
                ChDir Miruta
                Shell Miruta & "\end_day.bat"
            End If
        'Else
        '    ChDir MiruTa
        '    Shell MiruTa & "\findia.exe"
        'End If
    Else
        HabilitarSalir
        grventas
        lCogeTiquet = False
        '**************
        ' fin dia (exporta)
        '**************
        Trata_Fin_Dia
        If UsaBackup Then
            FrmResturarBackup.Realizar_Backup False
        End If
    End If
    lCogeTiquet = True
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Si Final_Dia=true se da la orden de borrado de totales, siempre y cuando, en configuración
' no se haya indicado lo contrario. También se hará copia de seguridad si procede
'///////////////////////////////////////////////////////////////////////////////////////////
Public Sub AN_RecogerTotales(Final_Dia As Boolean, Optional ByVal lNoMensaje)
Dim Reintentar As Boolean
Dim Contador As Integer
Dim Secciones As Typ_Tot_Sec
Dim ok_cont As Boolean
Dim bucle As Integer
Dim BucleEuroscale As Long
Dim Base As dao.Database
Dim Registro As dao.Recordset
Dim Descuadre As Boolean
Dim MiProceso As Long
Dim MiResultado As Long
'\\\\\\\\\\\\\\\\\\\\\\
Dim tipoTqt As Integer
Dim conxf As Integer
Dim sDirExportar As String
Dim Buf1 As String
Dim Buf2 As String
Dim sFields As String
Dim sfecha As String
Dim cnt As Long
Dim sSQL As String
Dim sext As String
Dim sDia As String
Dim sMes As String
Dim miArch As Integer
Dim OkBor As Long
Dim spa, cm As String
'/////////////////
Dim s As String

    On Error GoTo Sal_RecogerTotales
    canCelar = 0
    tipoTqt = 0
    CancelHumano = False
    'If Not Taquion Then 'c2f/caspiunza
        frmMonitor.Command1.Enabled = True
        frmMonitor.cmdmonitorstop.Enabled = True
    'End If
    
    '///
    'TouchScale
    '///
        
        lFinDia = Final_Dia
        nContDiasGi = 0
        
        tipoTqt = 1 '¿?
        
        MUE_Sistemas

        For BucleEuroscale = 0 To SisEur.NSistemas - 1
            
            CadenadeLog "An_RecogerTotales...Conexión"
            frmControl.Switch_mode 700 + BucleEuroscale
            
            'canCelar = 0
            If canCelar = 0 Then
            
                CancelHumano = False
                Contador = 0
                '''''''
                'Call capturaClientes
                ''''''''
                
                'quitar
                's = "update htickets set captured='' where status='C'"
                'MyConnObj.Execute s
                ''''
                
                LeeTickets 0
#If BALSAM = False Then
                If Dir(App.Path & "\nolabel.txt") = "" Then
                    If Dir(App.Path & "\nopre.txt") = "" Then
                        LeeTickets 1, Final_Dia
                    End If
                    If Dir(App.Path & "\nopack.txt") = "" Then
                        LeeTickets 2, Final_Dia
                    End If
                End If
#End If
                If Dir(App.Path & "\noinv.txt") = "" Then
                    If SolicitarPedido Then
                        If canCelar = 0 Then MUE_SolicitaPedido 0, False
                    End If
                    If SolicitarInventario Then
                        If canCelar = 0 Then MUE_SolicitaInventario 0, False
                    End If
                End If
                'touchScale Prueba...
                'CadenadeLog "Findia:" & Final_Dia & " GA_BorrarTotal:" & GA_BorrarTotal
                If Final_Dia And GA_BorrarTotal And canCelar = 0 Then
                    'CadenadeLog "1"
                    'If (Not Taquion) And (Not lNoMensaje) Then
                    '    frmMonitor.MostrarInfo CargaCadena(828)
                    'End If
                    'CadenadeLog "2"
                    
                    'TouchScale ... Realizar procedimiento Borrado Totales...
                    'CG_BorraTotales BucleEuroscale + 700
                     If GA_Reinit Then
                        Capa1A_GA_Erase_Totals (0)
                     Else
                        Capa1A_GA_Erase_Totals (1)
                     End If
                Else
                    If canCelar = 0 Then '''ver tienda
                        If TiendaActual = 0 Then
                            spa = App.Path & "\"
                        Else
                            spa = App.Path & "\t" & Format(TiendaActual, "00") & "\"
                        End If
                        miArch = FreeFile()
                        If Dir(spa & "ngtord") = "" Then
                            Open spa & "ngtord" For Output As #miArch
                            Print #miArch, "000001"
                            nGTOrd = 1
                        Else
                            Open spa & "ngtord" For Input As #miArch
                            Line Input #miArch, cm
                            Close #miArch
                            If Val(cm) > 9999 Then cm = "00000"
                            Open spa & "ngtord" For Output As #miArch
                            Print #miArch, Format(Val(cm) + 1, "000000")
                            nGTOrd = Val(cm) + 1
                        End If
                        Close #miArch
                    End If
                
                End If
                ''''''''''''''''''''''
                
                'On Error Resume Next
                'MyConnObj.Close
                'Set MyConnObj = Nothing
                'MyRecSet.Close
                'Set MyRecSet = Nothing
                'EstadoBal = 1000
                'On Error GoTo 0
                
                If CancelHumano Then Exit For
            
            Else
                CadenadeLog "Error en conexión..."
                HabilitarSalir
                Exit Sub
            End If
        Next BucleEuroscale
        
        'If HayMulti Then
        '    Compacta_PedidoInventario
        'End If
        
        CancelHumano = False
    
Sal_RecogerTotales:
    frmMonitor.Label1.Caption = CargaCadena(445)
    If Err.Number <> 0 Then
        CadenadeLog "Error en RecogerTotales --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
    End If
    On Error GoTo 0
    On Error GoTo sal
    
    'For OkBor = 0 To 9
        'Workspaces(OkBor).CommitTrans
        Workspaces(0).Close
    'Next OkBor
    
    HabilitarSalir
    'If Not Taquion Then
        frmMonitor.cmdmonitorstop.Enabled = False
    'End If
    '***********************************************
    ' compactar base de datos al final del día
    '***********************************************
    If Final_Dia Then
        'If Not Taquion Then
            frmMonitor.Command1.Enabled = False
        'End If
        Compactar_Base
        HabilitarSalir
    End If
    'If Dir(App.Path & "\alba.txt") <> "" Then
    If lAlba Then
        Alba_EnviaTiquetsPendientes (True)
    End If
    
sal:
    On Error GoTo 0
End Sub

Public Sub KILLSRV()
Dim Reintentar As Boolean
Dim Contador As Integer
Dim Secciones As Typ_Tot_Sec
Dim ok_cont As Boolean
Dim bucle As Integer
Dim BucleEuroscale As Long
Dim Base As dao.Database
Dim Registro As dao.Recordset
Dim Descuadre As Boolean
Dim MiProceso As Long
Dim MiResultado As Long
'\\\\\\\\\\\\\\\\\\\\\\
Dim tipoTqt As Integer
Dim conxf As Integer
Dim sDirExportar As String
Dim Buf1 As String
Dim Buf2 As String
Dim sFields As String
Dim sfecha As String
Dim cnt As Long
Dim sSQL As String
Dim sext As String
Dim sDia As String
Dim sMes As String
Dim miArch As Integer
Dim OkBor As Long
'/////////////////
    canCelar = 0
    tipoTqt = 0
    CancelHumano = False
    'If Not Taquion Then 'c2f/caspiunza
        frmMonitor.Command1.Enabled = True
        frmMonitor.cmdmonitorstop.Enabled = True
    'End If
    
    '///
    'TouchScale
    '///
        
        CadenadeLog "función KILLSRV..."
        MUE_Sistemas

        For BucleEuroscale = 0 To SisEur.NSistemas - 1
            
            frmControl.Switch_mode 700 + BucleEuroscale
            
            canCelar = 0
            CancelHumano = False
            Contador = 0
            '!!!!!!!!!!!!!!!!!!!!!!!!
            Call BorraSrv
            '!!!!!!!!!!!!!!!!!!!!!!!!
            
            On Error Resume Next
            MyConnObj.Close
            Set MyConnObj = Nothing
            MyRecSet.Close
            Set MyRecSet = Nothing
            EstadoBal = 1000
            On Error GoTo 0
            
            If CancelHumano Then Exit For
        
        Next BucleEuroscale
        
        CancelHumano = False
    
    
    HabilitarSalir
    'If Not Taquion Then
        frmMonitor.cmdmonitorstop.Enabled = False
    'End If
    
End Sub

Public Sub ExportarFicheroHIS(ByVal MiArchivo As String, ByVal MiFecha As Date, ByVal mIfECHA2 As Date)
    Dim Fichero As Integer
    Dim Buffer As String
    Dim FaEx As String
    MiTienda = ""
    Fichero = FreeFile
    FaEx = MiArchivo
    Open Miruta & "\etc\" & FaEx & ".fet" For Input As Fichero
    Line Input #Fichero, Buffer
    Close #Fichero
    If InStr(1, Buffer, "+") Then
        If HayMulti Then
            If TiendaActual < 100 Then
                MiTienda = Format(TiendaActual, "00")
            Else
                MiTienda = Format(TiendaActual, "000")
            End If
        End If
        Buffer = Mid(Buffer, InStr(1, Buffer, "+") - 1)
    End If
    If Not IsNumeric(Buffer) Then
        If MostrarMensajes Then MsgBox CargaCadena(848), vbCritical, vbOKOnly
        Exit Sub
    End If
    Select Case Val(Buffer)
        Case 1
            ExportaGTarti Exp_Path & "\" & FaEx, MiFecha, mIfECHA2
        Case 2
            ExportaGTSecs Exp_Path & "\" & FaEx, MiFecha, mIfECHA2
        Case 3
            ExportaGTVend Exp_Path & "\" & FaEx, MiFecha, mIfECHA2
        Case 4
            ExportaCabecera Exp_Path & "\" & FaEx, MiFecha, False, mIfECHA2
        Case 5
            ExportaLinea Exp_Path & "\" & FaEx, MiFecha, False, mIfECHA2
        Case 6
            ExportaCabLin Exp_Path & "\" & FaEx, MiFecha, False, mIfECHA2
        Case Else
            If MostrarMensajes Then MsgBox CargaCadena(848), vbCritical, vbOKOnly
            Exit Sub
    End Select

End Sub





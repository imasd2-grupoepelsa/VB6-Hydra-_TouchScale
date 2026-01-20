VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMantenimiento 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANTENIMIENTO"
   ClientHeight    =   3495
   ClientLeft      =   7155
   ClientTop       =   2235
   ClientWidth     =   4500
   Icon            =   "mantenimiento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Opttotal 
      Appearance      =   0  'Flat
      Caption         =   "Borrar total diario"
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
   Begin VB.OptionButton Opttotal 
      Appearance      =   0  'Flat
      Caption         =   "Borrar total acumulado"
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
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2532
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Periodo de borrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1692
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4212
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   5
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   14
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   4
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   13
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   3
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   12
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   2
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   11
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   1
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   10
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox TxtFecha 
         Appearance      =   0  'Flat
         Height          =   372
         Index           =   0
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   9
         Top             =   360
         Width           =   372
      End
      Begin MSMask.MaskEdBox TxtDesde 
         Height          =   372
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDesde 
         Height          =   372
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   3120
         TabIndex        =   18
         Top             =   960
         Width           =   372
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   2400
         TabIndex        =   17
         Top             =   960
         Width           =   372
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   372
      End
      Begin VB.Label lbldesde 
         Caption         =   "Borrar desde fecha:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label lbldesde 
         Caption         =   "Borrar hasta fecha:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1692
      End
   End
   Begin VB.CommandButton cmdespere 
      Caption         =   "Aceptar"
      Default         =   -1  'True
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
      TabIndex        =   4
      Top             =   3000
      Width           =   1452
   End
   Begin VB.CommandButton cmdespere_terminar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'\\\\\\\\\\\\\\
'//////////////
Public MostrarMsg As Boolean
'\\\\\\\\\\\\\\
'//////////////
Private Sub Form_Load()
Dim Fecha As Date
Dim bucle As Integer
Dim pepe As String
    MostrarMsg = True
    Me.Icon = Form2.Icon
    If Dir(App.Path & "\etc\l_date.txt") <> "" Then
        For bucle = 0 To 5
            TxtFecha(bucle).Visible = True
        Next bucle
        For bucle = 0 To 3
            Label1(bucle).Visible = True
        Next bucle
        TxtDesde(0).Visible = False
        TxtDesde(1).Visible = False
    Else
        For bucle = 0 To 5
            TxtFecha(bucle).Visible = False
        Next bucle
        TxtDesde(0).Visible = True
        TxtDesde(1).Visible = True
    End If
    frmMantenimiento.Caption = UCase(CargaCadena(455))
    Frame3.Caption = CargaCadena(461)
    Opttotal(0).Caption = CargaCadena(239)
    Opttotal(1).Caption = CargaCadena(124)
    lbldesde(0).Caption = CargaCadena(457)
    lbldesde(1).Caption = CargaCadena(458)
    Fecha = Now()
    TxtDesde(0).TexT = AdaptaFecha(Format(Fecha, "dd/MM/yy"))
    TxtDesde(1).TexT = AdaptaFecha(Format(Fecha, "DD/MM/YY"))
    Opttotal(0).Value = True
    cmdespere.Caption = CargaCadena(287)
    cmdespere_terminar.Caption = CargaCadena(288)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub
'\\\\\\\\\\\\\\
'//////////////
Public Sub cmdespere_Click()
Dim Base As dao.Database
Dim fecha_ini As Date
Dim fecha_fin As Date
Dim Fecha As Date
Dim rege_sino
Dim Registro As dao.Recordset
Dim bucle As Integer
Dim Max As Integer
Dim MiTabla As String
Dim MiNombre As String
Dim lEsPrimera As Boolean

lEsPrimera = False

    If Dir(App.Path & "\borrados", vbDirectory) = "" Then
        MkDir App.Path & "\borrados"
        Sleep (100)
    End If
    
    If frmMantenimiento.Visible = True Then
        FileCopy Base_General, App.Path & "\borrados\" & Format(TiendaActual, "000") & Format(Now, "ddmmyyhhmmss") & "dbase.mdb"
    End If
    
    '3.4.4
    'If lNoSi = False Then
    '    Interfaz.exe_copion "procsiosi"
    'End If
    ''''''
    
    If Opttotal(0).Value = True Then
        If Dir(App.Path & "\etc\l_date.txt") = "" Then
            If (Not IsNumeric(left(TxtDesde(0).TexT, 2))) Or _
            (Not IsNumeric(Right(TxtDesde(0).TexT, 2))) Or _
            (Not IsNumeric(Mid(TxtDesde(0).TexT, 4, 2))) Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If (Not IsNumeric(left(TxtDesde(1).TexT, 2))) Or _
            (Not IsNumeric(Right(TxtDesde(1).TexT, 2))) Or _
            (Not IsNumeric(Mid(TxtDesde(1).TexT, 4, 2))) Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
        Else
            If Val(TxtFecha(0)) < 1 Or Val(TxtFecha(1)) > 31 Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If Val(TxtFecha(3)) < 1 Or Val(TxtFecha(3)) > 31 Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If Val(TxtFecha(1)) < 1 Or Val(TxtFecha(1)) > 12 Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If Val(TxtFecha(4)) < 1 Or Val(TxtFecha(4)) > 12 Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If Not IsNumeric(TxtFecha(2).TexT) Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
            If Not IsNumeric(TxtFecha(5).TexT) Then
                If MostrarMsg Then MsgBox CargaCadena(506), vbCritical
                Exit Sub
            End If
        End If
    End If
    'c2f invicta parámetro arranque borrado
    'If MostrarMsg Then
    '    rege_sino = MsgBox(CargaCadena(145) & "?", vbYesNo)
    'Else
    '    rege_sino = vbYes
    'End If
    '...no se pide confirmación
    rege_sino = vbYes
    '''''''''''''''''''''''''''
    If rege_sino = vbYes Then
    
    'c2f prueba...
    'CadenadeLog "HOLA INVICTA"
    ''''''''''''''
        If Trim(Command) = "" Then
            Load frmWAIT
            frmWAIT.Label1.Caption = "Eliminando totales antiguos..."
            frmWAIT.Visible = True
            frmWAIT.Refresh
        End If
        Do_Events
        
        Set Base = OpenDatabase(Base_General)
        fecha_ini = DateSerial(Right(TxtDesde(0).TexT, 2), Mid(TxtDesde(0).TexT, 4, 2), left(TxtDesde(0).TexT, 2))
        fecha_fin = DateSerial(Right(TxtDesde(1).TexT, 2), Mid(TxtDesde(1).TexT, 4, 2), left(TxtDesde(1).TexT, 2))
        If UsaFichasVacuno Then
            Max = 14
        Else
            Max = 13
        End If
        For bucle = 0 To Max
            Select Case bucle
                Case 0
                    MiTabla = "gtsecs"
                Case 1
                    MiTabla = "gtarti"
                Case 2
                    MiTabla = "gtvend"
                Case 3, 4, 7
                    MiTabla = "tickets"
                'Case 4
                '    MiTabla = "ticketsSC10"
                Case 5, 6, 8
                    MiTabla = "cabecera"
                'Case 6
                '    MiTabla = "cabeceraSC10"
                'Case 7
                '    MiTabla = "ticketsTQ"
                'Case 8
                '    MiTabla = "cabeceraTQ"
                Case 9
                    MiTabla = "log"
                Case 10
                    MiTabla = "rechazos"
                Case 11
                    MiTabla = "gtInfVen"
                Case 12
                    MiTabla = "descuentos"
                Case 13
                    MiTabla = "gthora"
                Case 14
                    MiTabla = "gtvacuno"
            End Select
            
            If Trim(Command) = "" Then
                frmWAIT.Label2.Caption = "Tabla: " & MiTabla
                frmWAIT.Refresh
            End If
            Do_Events
            
            Set Registro = Base.OpenRecordset("select * from " & MiTabla)
            With Registro
                Do While Not .EOF
                    If Opttotal(0).Value = True And MiTabla <> "rechazos" Then
                    '    Fecha = Format(.Fields(1).Value, "ddmmyy")
                    '    Fecha = Mid(Fecha, 5, 2) & Mid(Fecha, 3, 2) & Mid(Fecha, 1, 2)
                        If (bucle < 3) Or (bucle = 7) Or ((bucle >= 11) And (bucle <= 13)) Then
                            Select Case bucle
                                Case 0 To 2
                                    MiNombre = "fecha"
                                Case 3 To 4
                                    MiNombre = "date"
                                Case 5 To 6
                                    MiNombre = "fecha"
                                Case 7
                                    MiNombre = "date"
                                Case 8
                                    MiNombre = "fecha"
                                Case 9
                                    MiNombre = "fecha"
                                Case 11 To 13
                                    MiNombre = "fecha"
                                Case 14
                                    MiNombre = "fecha"
                            End Select
                            Fecha = .Fields(MiNombre).Value
                        Else
                            If bucle < 5 Then
                                Fecha = DateSerial(Right(!Date, 2), Mid(!Date, 4, 2), left(!Date, 2))
                            Else
                                'c2f/caspiunza ** If Bucle < 9 Or Bucle = 11 Then
                                If bucle < 9 Then
                                    Fecha = DateSerial(Right(!Fecha, 2), Mid(!Fecha, 4, 2), left(!Fecha, 2))
                                Else
                                    Fecha = .Fields("fecha").Value
                                
                                End If
                                
                            End If
                        
                        End If
                        If Fecha >= fecha_ini And Fecha <= fecha_fin Then
                            'c2f copias a ficheros formato totales.dat si se van a borrar
                            '    y es multitienda.
                            If Dir(App.Path & "\hydramulti.exe") <> "" And lEsPrimera = False Then
                               lEsPrimera = True
                               ExportaTotalesDat True
                            End If
                            If Not .EOF Then Delete_Record Registro
                        End If
                    Else
                        If Not .EOF Then Delete_Record Registro
                    End If
                    If Not .EOF Then .Movenext
                Loop
                .Close
            End With
       Next bucle
       'If Trim(Command) = "" Then '2.0.26
           CerrarBase Base
       'End If
       'c2f invicta no confirmación por parámetro
       'If MostrarMsg Then MsgBox "OK"
       ''''''''''''''''''''''''''''''''''''''''''
       If Trim(Command) = "" Then
           Unload frmWAIT
           Set frmWAIT = Nothing
       End If
       
       grventas
   End If
    cmdespere_terminar_Click
End Sub
Private Sub cmdespere_terminar_Click()
    Unload Me
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub Opttotal_Click(Index As Integer)
    If Index = 1 Then
        Frame3.Enabled = False
    ElseIf Index = 0 Then
        Frame3.Enabled = True
    End If
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub TxtFecha_Change(Index As Integer)
    If Index < 5 Then
        If Len(TxtFecha(Index).TexT) = 2 Then
            On Error Resume Next
            TxtFecha(Index + 1).SetFocus
            On Error GoTo 0
        End If
    End If
    Select Case Index
        Case 0 To 2
            TxtDesde(0).TexT = Format(Val(TxtFecha(0).TexT), "00") & "/" & Format(Val(TxtFecha(1).TexT), "00") & "/" & Format(Val(TxtFecha(2).TexT), "00")
        Case 3 To 5
            TxtDesde(1).TexT = Format(Val(TxtFecha(3).TexT), "00") & "/" & Format(Val(TxtFecha(4).TexT), "00") & "/" & Format(Val(TxtFecha(5).TexT), "00")
    End Select
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub TxtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

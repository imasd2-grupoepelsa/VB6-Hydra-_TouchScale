VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form fechaHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FECHA HISTORICO"
   ClientHeight    =   8520
   ClientLeft      =   7830
   ClientTop       =   420
   ClientWidth     =   3240
   Icon            =   "fechaHistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdMarcar 
      Caption         =   "Marcar en Balanza."
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox TxtNum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox TxtName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "Restaurar en  Dbase.Mdb"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   3255
   End
   Begin MSComctlLib.TreeView tvwBck 
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7435
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton cmdth 
      Caption         =   "Tiques"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton cmdBck 
      Caption         =   "INFORME"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   3255
   End
   Begin VB.FileListBox flbBck 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3195
   End
   Begin MSMask.MaskEdBox MFF 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   7440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MFI 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Lbl1 
      Caption         =   "Marcado en Base Balanzas para Captura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "F.Inicial:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "F.Final:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   7440
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   6240
      Y2              =   6240
   End
End
Attribute VB_Name = "fechaHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdth_Click()
    Dim sruta As String
    Dim nFich As Integer
    If ((Me.flbBck.ListCount > 0) And (Me.tvwBck.SelectedItem.Index <> 1)) Then
        Me.Enabled = False
        sruta = Me.flbBck.Path & "\" & Right(Me.tvwBck.SelectedItem.Key, 6) & ".mdb"
        'aux_ejecuta_historico "gtsecs", "0", sruta
        'Unload Me
        sruta = Replace(sruta, "..", ".")
        nFich = FreeFile()
        Open App.Path & "\ht.txt" For Output As nFich
        Print #nFich, sruta
        Close #nFich
        Load FrmConsultaTiquets
        FrmConsultaTiquets.Show (vbModal)
        Me.Enabled = True
    End If

End Sub

Private Sub CmdMarcar_Click()
    Dim Resp As Integer
    Dim cm As String

    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox "No tiene permiso para realizar esta operación...", vbCritical
        Exit Sub
    End If
    If IsDate(MFI.TexT) And IsDate(MFF.TexT) Then
        Resp = MsgBox("Atención se van a marcar en Balanza los Tiques para ser capturados de nuevo...", vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbYes Then
            Resp = MsgBox("Confirme de nuevo esta acción...", vbOKCancel + vbDefaultButton2 + vbExclamation)
            If Resp = vbOK Then
                'Call Absorver_Estructura
                frmControl.Switch_mode 0
                If EstadoBal = 0 Then
                    If CDate(MFI.TexT) <= CDate(MFF.TexT) Then
                        cm = "update pcscale.htickets set captured='',CapturedDateTime='1970-01-01 00:00:00' where substr(dattim,1,10) >= '" & Format(MFI.TexT, "yyyy-mm-dd") & "' and substr(dattim,1,10) <= '" & Format(MFF.TexT, "yyyy-mm-dd") & "'"
                    Else
                        cm = "update pcscale.htickets set captured='',CapturedDateTime='1970-01-01 00:00:00' where substr(dattim,1,10) >= '" & Format(MFF.TexT, "yyyy-mm-dd") & "' and substr(dattim,1,10) <= '" & Format(MFI.TexT, "yyyy-mm-dd") & "'"
                    End If
                    On Error Resume Next
                    MyConnObj.Execute cm
                    If Err.Number = 0 Then
                        MsgBox "Tiques marcados en balanza. Reciba tiques...", vbInformation
                    Else
                        On Error GoTo 0
                        MsgBox "No ha sido posible el marcado de tiques para su recepción...", vbExclamation
                    End If
                    On Error GoTo 0
                Else
                    MsgBox "No hay conexión con Balanza...", vbExclamation
                End If
            End If
        End If
    Else
        MsgBox "Fechas erroneas...", vbExclamation
    End If

End Sub

Private Sub cmdrestore_Click()
    Dim sruta As String
    Dim sRutaBase As String
    Dim nFich As Integer
    Dim sBaseOri As dao.Database
    Dim sBaseDes As dao.Database
    Dim sTablas(100) As String
    Dim nCampos(100) As Integer
    Dim sCampos(100) As String
    Dim nContTablas As Integer
    Dim nContFields As Integer
    Dim n As Integer
    Dim m As Integer
    Dim nOriR As dao.Recordset
    Dim nDesR As dao.Recordset
    Dim nCONTX As Double
    Dim nPp As Integer

    If ((Me.flbBck.ListCount > 0) And (Me.tvwBck.SelectedItem.Index <> 1)) Then
        TxtName.Visible = True
        TxtNum.Visible = True
        Me.Enabled = False
        sruta = Me.flbBck.Path & "\" & Right(Me.tvwBck.SelectedItem.Key, 6) & ".mdb"
        sruta = UCase(sruta)

        'aux_ejecuta_historico "gtsecs", "0", sruta
        'Unload Me
        'nFich = FreeFile()
        'nFich = InStr(1, "\S\", sruta)

        nFich = InStr(1, sruta, "\S\")
        If nFich <> 0 Then
            sRutaBase = Mid(sruta, 1, nFich - 1) & "\dbasetouch.mdb"
            On Error Resume Next
            Set sBaseOri = OpenDatabase(sruta)
            If Err.Number = 0 Then
                Set sBaseDes = OpenDatabase(sRutaBase)
                If Err.Number = 0 Then
                    nContTablas = sBaseOri.TableDefs.Count
                    m = 0
                    For n = 0 To nContTablas - 1
                        If (UCase(Mid(sBaseOri.TableDefs(n).name, 1, 4)) <> "MSYS") And (UCase(Mid(sBaseOri.TableDefs(n).name, 1, 7)) <> "CLIENTE") Then
                            m = m + 1
                            sTablas(m) = sBaseOri.TableDefs(n).name
                            nCampos(m) = sBaseOri.TableDefs(n).Fields.Count
                        End If
                    Next n
                    nContTablas = m
                    m = 0
                    For n = 1 To nContTablas
                        nPp = 1
                        CadenadeLog "TABLA:" & "N.:" & CStr(n) & " -->" & sTablas(n)
                        Set nOriR = sBaseOri.OpenRecordset("select * from " & sTablas(n))
                        Set nDesR = sBaseDes.OpenRecordset("select * from " & sTablas(n))

                        If Not nOriR.EOF Then
                            nPp = 2
                            nOriR.MoveFirst

                            TxtName.TexT = sTablas(n)
                            nCONTX = 0

                            Do While Not nOriR.EOF
                                nPp = 3
                                nDesR.AddNew
                                nCONTX = nCONTX + 1
                                For m = 0 To nCampos(n) - 1
                                    'If LCase(nOriR.Fields(m).name) = "factura" Then
                                    '    m = m
                                    'End If
                                    nDesR.Fields(nOriR.Fields(m).name) = nOriR.Fields(nOriR.Fields(m).name)
                                    'nDesR.Fields(m).name = nOriR.Fields(m).name
                                Next m
                                nDesR.Update
                                nOriR.Delete
                                TxtNum.TexT = CStr(nCONTX)

                                'Do_Events
                                'Sleep (1)
                                TxtName.Refresh
                                TxtNum.Refresh
                                ''''''''''
                                If Not nOriR.EOF Then
                                    nOriR.MoveFirst
                                Else
                                    Exit Do
                                End If
                            Loop

                            nDesR.Close

                            Set nDesR = Nothing

                        End If

                        nOriR.Close
                        Set nOriR = Nothing

                    Next n
                    nPp = 4
                    'If Err.Number = 0 Then
                    MsgBox "Restore OK!", vbInformation
                    'End If

                End If
                nPp = 5
                sBaseDes.Close
                Set sBaseDes = Nothing
                nPp = 6
            End If
            sBaseOri.Close
            Set sBaseOri = Nothing
            If Err.Number <> 0 Then
                CadenadeLog "No Restore..." & CStr(Err.Number) & " --> " & Err.Description
                CadenadeLog "Paso:" & nPp
            End If
            On Error GoTo 0
        Else
            MsgBox "No Restore...", vbExclamation
        End If
        Me.Enabled = True
        TxtName.Visible = False
        TxtNum.Visible = False

    End If

    'Corregir_Nulos

End Sub


'''''''''''''''
'''''''''''''''
Private Sub CambiarIdioma()
    fechaHistorico.Caption = CargaCadena(1355)
    cmdBck.Caption = CargaCadena(1352)
    cmdth.Caption = CargaCadena(1353)
    CmdRestore.Caption = CargaCadena(1354)

End Sub
Private Sub Form_Load()
    Dim cnt As Long
    Dim Key As String
    Dim RelativE As String
    Dim RelationShiP As Long
    Dim mm As String
    Dim syyyy As String
    Dim TexT As String
    TxtName.Visible = False
    TxtNum.Visible = False
    Call CambiarIdioma
    With Me
        If HayMulti Then
            If TiendaActual < 100 Then
                .flbBck.Path = App.Path & "\T" & Format(TiendaActual, "00") & "\S"
            Else
                .flbBck.Path = App.Path & "\T" & Format(TiendaActual, "000") & "\S"
            End If
        Else
            .flbBck.Path = App.Path & "\S"
        End If
        .tvwBck.Nodes.Clear
        Key = "BCK"
        TexT = CargaCadena(1356)    '"BASE DATOS HISTORICO"
        .tvwBck.Nodes.Add , , Key, TexT
        RelativE = Key
        RelationShiP = tvwChild
        For cnt = 0 To .flbBck.ListCount - 1
            If Len(.flbBck.List(cnt)) > 4 Then
                mm = CStr(Mid(.flbBck.List(cnt), 5, 2))
                syyyy = CStr(left(.flbBck.List(cnt), 4))
                Key = RelativE & syyyy & mm
                Select Case Val(mm)
                Case 1
                    'TexT = "ENERO" & " - " & syyyy
                    TexT = CargaCadena(1357) & " - " & syyyy
                Case 2
                    TexT = CargaCadena(1358) & " - " & syyyy
                Case 3
                    TexT = CargaCadena(1359) & " - " & syyyy
                Case 4
                    TexT = CargaCadena(1360) & " - " & syyyy
                Case 5
                    TexT = CargaCadena(1361) & " - " & syyyy
                Case 6
                    TexT = CargaCadena(1362) & " - " & syyyy
                Case 7
                    TexT = CargaCadena(1363) & " - " & syyyy
                Case 8
                    TexT = CargaCadena(1364) & " - " & syyyy
                Case 9
                    TexT = CargaCadena(1365) & " - " & syyyy
                Case 10
                    TexT = CargaCadena(1366) & " - " & syyyy
                Case 11
                    TexT = CargaCadena(1367) & " - " & syyyy
                Case 12
                    TexT = CargaCadena(1368) & " - " & syyyy
                End Select
                On Error Resume Next
                .tvwBck.Nodes.Add RelativE, RelationShiP, Key, TexT
                On Error GoTo 0
            End If
        Next cnt
        .tvwBck.Nodes(1).Expanded = True
    End With
End Sub
'''''''
'''''''
Private Sub cmdBck_Click()
    Dim sruta As String
    If ((Me.flbBck.ListCount > 0) And (Me.tvwBck.SelectedItem.Index <> 1)) Then
        Me.Enabled = False
        sruta = Me.flbBck.Path & "\" & Right(Me.tvwBck.SelectedItem.Key, 6) & ".mdb"
        aux_ejecuta_historico "gtsecs", "0", sruta
        Unload Me
        Me.Enabled = True
    End If
End Sub
'''''''
'''''''

Private Sub MFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdMarcar.SetFocus
End Sub

Private Sub MFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MFF.SetFocus
End Sub

Private Sub MFI_LostFocus()
    MFF.SetFocus
End Sub

Private Sub tvwBck_Click()
    tvwBck.SetFocus
End Sub

VERSION 5.00
Begin VB.Form FrmGesco 
   Caption         =   "Generación Ficheros datos GESCOWIN"
   ClientHeight    =   2580
   ClientLeft      =   4440
   ClientTop       =   675
   ClientWidth     =   7200
   Icon            =   "FrmGesco.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7200
   Begin VB.Timer TimerGesco 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2040
   End
   Begin VB.TextBox TxtNumReg 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtTienda 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtDate 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label LblWait 
      Caption         =   "POR FAVOR ESPERE..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label LblNumReg 
      Caption         =   "Número de Registros:"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label LblTienda 
      Caption         =   "Tienda en Curso:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label LblDate 
      Caption         =   "Generación para Fecha:"
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
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblinfo 
      Caption         =   $"FrmGesco.frx":1E72
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "FrmGesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TimerGesco.Enabled = True
    CadenadeLog "Entrada en exportación Ficheros GescoWin"
End Sub

Private Sub GeneraGesco()
    Dim BaseT As dao.Database
    Dim Tiendas As dao.Database
    Dim rB As dao.Recordset
    Dim rT As dao.Recordset
    Dim sCabecera As String
    Dim sDetalle As String
    Dim nFGesco As Integer
    Dim sPathB As String
    Dim Supertot As Double
    Dim nFNum As Integer
    Dim sFNum As String
    Dim sfecha As String
    Dim ncont As Long
    Dim sGes As String
    Dim sP As String
    Dim nnP As Double

    sP = ""
    If Dir(App.Path & "\gescopath.txt") <> "" Then
        nFNum = FreeFile()
        Open App.Path & "\gescopath.txt" For Input As #nFNum
        Line Input #nFNum, sP
        Close #nFNum
        sP = Trim(sP)
        If Right(sP, 1) <> "\" Then
            sP = sP & "\"
        End If
    Else
        sP = Trim(App.Path) & "\"
    End If

    sfecha = sFechaGesco

    Dim lHay As Boolean

    LblWait.BackColor = vbGreen
    Me.Refresh

    txtDate.TexT = sfecha
    lHay = False

    If Dir(App.Path & "\tiendas.mdb") <> "" And Dir(App.Path & "\gesconum.txt") <> "" Then

        Set Tiendas = OpenDatabase(App.Path & "\tiendas.mdb")
        Set rT = Tiendas.OpenRecordset("select * from grupos where tienda<>0")
        If Not rT.EOF Then

            rT.MoveFirst
            LblWait.BackColor = vbBlue
            Me.Refresh

            nFNum = FreeFile()
            Open App.Path & "\gesconum.txt" For Input As #nFNum
            Line Input #nFNum, sFNum
            Close #nFNum

            Do While Not rT.EOF

                sPathB = rT.Fields("path") & "dbasetouch.mdb"

                If Dir(sPathB) <> "" Then

                    sGes = sP & "datos3" & CStr(rT.Fields("tienda"))

                    If Dir(sGes & ".dat") <> "" Then
                        FileCopy sGes & ".dat", sGes & ".bak"
                    End If
                    nFGesco = FreeFile()
                    Open sGes & ".dat" For Output As #nFGesco

                    txtTienda.TexT = CStr(rT.Fields("tienda"))
                    LblWait.BackColor = vbGreen
                    Me.Refresh

                    CadenadeLog sGes & " Tienda:" & CStr(rT.Fields("tienda"))
                    Set BaseT = OpenDatabase(sPathB)
                    Set rB = BaseT.OpenRecordset("select sum(pesetas) as total_imp from gtarti where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(sfecha)) & Chr(34))
                    If rB.Fields("total_imp") <> 0 Then

                        lHay = True

                        LblWait.BackColor = vbBlue
                        Me.Refresh

                        Supertot = rB.Fields("total_imp")
                        rB.Close

                        sCabecera = "3" & CStr(rT.Fields("tienda")) & "3" & CStr(rT.Fields("tienda")) & "1" & "T1"
                        sCabecera = sCabecera & Format(Val(sFNum), "000000") & "5" & "000902" & CStr(rT.Fields("tienda"))
                        sCabecera = sCabecera & Format(CStr(CDate(sfecha)), "DDMMYYYY") & Format(Time, "HH:MM")
                        sCabecera = sCabecera & Format(Supertot * 100, "00000000") & "00000.00" & "00.00" & "A"
                        sCabecera = sCabecera & "000" & "00000000" & "N" & "0000000000"

                        Print #nFGesco, sCabecera

                        Set rB = Nothing
                        Set rB = BaseT.OpenRecordset("select * from gtarti where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(sfecha)) & Chr(34) & " order by codigo")
                        If Not rB.EOF Then
                            rB.MoveFirst
                            ncont = 0
                            Do While Not rB.EOF

                                LblWait.BackColor = vbGreen
                                ncont = ncont + 1
                                TxtNumReg.TexT = CStr(ncont)

                                sDetalle = "3" & CStr(rT.Fields("tienda")) & "3" & CStr(rT.Fields("tienda")) & "2" & "T1"
                                sDetalle = sDetalle & Format(Val(sFNum), "000000") & Format(rB.Fields("codigo"), "000000")
                                If rB.Fields("peso") <> 0 Then
                                    If rB.Fields("peso") >= 0 Then
                                        sDetalle = sDetalle & Format(rB.Fields("peso") * 100, "0000000")
                                        If rB.Fields("peso") > 0 Then
                                            nnP = (rB.Fields("pesetas")) / (rB.Fields("peso"))
                                        Else
                                            nnP = 0
                                        End If
                                    Else
                                        sDetalle = sDetalle & Format(rB.Fields("peso") * 100, "000000")
                                        nnP = 0
                                    End If


                                Else
                                    If rB.Fields("unidades") >= 0 Then
                                        sDetalle = sDetalle & Format(rB.Fields("unidades") * 100, "0000000")
                                        If rB.Fields("unidades") > 0 Then
                                            nnP = (rB.Fields("pesetas")) / (rB.Fields("unidades"))
                                        Else
                                            nnP = 0
                                        End If
                                    Else
                                        sDetalle = sDetalle & Format(rB.Fields("unidades") * 100, "000000")
                                        nnP = 0
                                    End If

                                End If
                                'se debe poner precio medio=nnp  sDetalle = sDetalle & "0000000" & Format(rB.Fields("precio"), "0000.00")
                                If Dir(App.Path & "\080418.txt") <> "" Then
                                    If nnP <> 0 And rB.Fields("tipo_iva") <> 0 Then    ' precio sin IVA
                                        Select Case rB.Fields("tipo_iva")
                                        Case 1
                                            nnP = nnP / (1 + 0.08)
                                        Case 2
                                            nnP = nnP / (1 + 0.04)
                                        Case 3
                                            nnP = nnP / (1 + 0.18)
                                        End Select
                                    End If
                                Else
                                    If nnP <> 0 And rB.Fields("tipo_iva") <> 0 Then    ' precio sin IVA
                                        Select Case rB.Fields("tipo_iva")
                                        Case 1
                                            nnP = nnP / (1 + 0.07)
                                        Case 2
                                            nnP = nnP / (1 + 0.04)
                                        Case 3
                                            nnP = nnP / (1 + 0.16)
                                        End Select
                                    End If
                                End If

                                sDetalle = sDetalle & "0000000" & Format(nnP, "0000.00")

                                If rB.Fields("pesetas") > 0 Then
                                    sDetalle = sDetalle & Format(rB.Fields("pesetas"), "00000.00")
                                Else
                                    sDetalle = sDetalle & Format(rB.Fields("pesetas"), "0000.00")
                                End If

                                sDetalle = sDetalle & "0000" & "V" & "A"

                                Print #nFGesco, Quitar_Coma(sDetalle, Len(sDetalle))


                                rB.Movenext
                                LblWait.BackColor = vbBlue
                                Me.Refresh

                                ncont = ncont + 1
                                TxtNumReg.TexT = CStr(ncont)

                                Do_Events

                            Loop



                        End If


                    End If

                    Close #nFGesco
                    If Dir(App.Path & "\gescoincre.txt") <> "" Then
                        sFNum = Format(Val(sFNum) + 1, "000000")
                    End If

                    rB.Close
                    Set rB = Nothing
                    BaseT.Close
                    Set BaseT = Nothing


                End If

                rT.Movenext
            Loop



            If lHay Then
                nFNum = FreeFile()
                Open App.Path & "\gesconum.txt" For Output As #nFNum
                If Dir(App.Path & "\gescoincre.txt") <> "" Then
                    Print #nFNum, Format(Val(sFNum), "000000")
                Else
                    Print #nFNum, Format(Val(sFNum) + 1, "000000")
                End If
                Close #nFNum
            End If

        End If

        rT.Close
        Set rT = Nothing
        Tiendas.Close
        Set Tiendas = Nothing

    End If

    nFGesco = FreeFile()
    Open App.Path & "\gescofin" For Output As #nFGesco
    Close #nFGesco

    Unload Me

End Sub

Private Sub TimerGesco_Timer()
    TimerGesco.Enabled = False
    CadenadeLog "LLamada función exportación..."
    Call GeneraGesco

End Sub

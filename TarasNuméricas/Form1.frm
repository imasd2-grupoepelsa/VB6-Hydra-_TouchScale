VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmTaras 
   Caption         =   "Programación Taras Numéricas."
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImpTaras 
      BackColor       =   &H8000000E&
      Caption         =   "Importar Taras.Dat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox ChkSvr 
      Caption         =   "Enviar a Servidora."
      Height          =   315
      Left            =   3840
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtval 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   9
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtNum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox txtBal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   5
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton CmdEnvTodas 
      BackColor       =   &H8000000E&
      Caption         =   "Enviar a Todas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton CmdTaras 
      BackColor       =   &H8000000E&
      Caption         =   "Modificar Taras Numéricas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGridEqp 
      Bindings        =   "Form1.frx":1E72
      Height          =   5895
      Left            =   0
      OleObjectBlob   =   "Form1.frx":1E88
      TabIndex        =   0
      Top             =   840
      Width           =   8535
   End
   Begin VB.Data DataEqp 
      Connect         =   "Access"
      DatabaseName    =   "C:\hydra+1720\TarasNuméricas\dbase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "equipos"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Valor:"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "N.Tara:"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Balanza:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   6960
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   480
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3480
      Y1              =   840
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Balanzas de la Tienda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "FrmTaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ChkSvr_Click()
    If ChkSvr.Value = vbChecked Then
        CmdEnvTodas.Caption = "Enviar a Servidora."
        DataEqp.DatabaseName = App.Path & "\dbase.mdb"
        DataEqp.RecordSource = "select * from equipos where modelo=1"
        DBGridEqp.Refresh
    Else
        CmdEnvTodas.Caption = "Enviar a Todas."
        DataEqp.DatabaseName = App.Path & "\dbase.mdb"
        DataEqp.RecordSource = "select * from equipos"
        DataEqp.Refresh
        DBGridEqp.Refresh
    
    End If
End Sub

Private Sub CmdEnvTodas_Click()
    Dim Base As Database
    Dim regEqp As Recordset
    Dim regTar As Recordset
    Dim lC As Boolean
    Dim nR As Long
    
    CmdTaras.Enabled = False
    CmdEnvTodas.Enabled = False
    txtBal.Text = ""
    txtNum.Text = ""
    txtval.Text = ""
    
    Set Base = OpenDatabase(App.Path & "\dbase.mdb", False, False)
    If ChkSvr.Value = vbChecked Then
        Set regEqp = Base.OpenRecordset("select * from equipos where modelo=1")
    Else
        Set regEqp = Base.OpenRecordset("select * from equipos")
    End If
    Set regTar = Base.OpenRecordset("select * from tipotara")
    regEqp.MoveFirst
    Do While Not regEqp.EOF
        txtBal.Text = CStr(regEqp.Fields("numero_eqp"))
        
        sIp = Val(Left(regEqp.Fields("prog_eqp"), 3)) & "." & Val(Mid(regEqp.Fields("prog_eqp"), 5, 3)) & "." & Val(Mid(regEqp.Fields("prog_eqp"), 9, 3)) & "." & Val(Right(regEqp.Fields("prog_eqp"), 3))
        CadenadeLog "Balanza:" & txtBal.Text & " Dirección IP:" & sIp & "Port: 6000"
        
        lC = abreSOCKETuno()
        
        If lC Then 'se conecta equipo
            CadenadeLog "Socket Conectado..."
            regTar.MoveFirst
            Do While Not regTar.EOF
                txtNum.Text = CStr(regTar.Fields("tipo"))
                txtval.Text = CStr(regTar.Fields("valor"))
                nR = SendTare(regTar.Fields("tipo"), regTar.Fields("valor"), regEqp.Fields("numero_eqp"))
                regTar.MoveNext
            Loop
        End If
        
        lC = cierraSOCKETuno()
        
        If lC Then
            CadenadeLog "Socket Cerrado"
        End If
        regEqp.MoveNext
    Loop
    regTar.Close
    regEqp.Close
    Base.Close
    Set regTar = Nothing
    Set regEqp = Nothing
    Set Base = Nothing
    
    CmdTaras.Enabled = True
    CmdEnvTodas.Enabled = True
    
    txtBal.Text = ""
    txtNum.Text = ""
    txtval.Text = ""

End Sub

Private Sub CmdTaras_Click()
    CmdTaras.Enabled = False
    CmdEnvTodas.Enabled = False
    
    FrmTipoTARAGA.Show vbModal

    CmdTaras.Enabled = True
    CmdEnvTodas.Enabled = True

End Sub


'With registro
'    If Not .EOF Then
'        If Len(!prog_eqp) = 15 Then
'            If IsNumeric(Left(!prog_eqp, 3)) And IsNumeric(Mid(!prog_eqp, 5, 3)) And IsNumeric(Mid(!prog_eqp, 9, 3)) And IsNumeric(Right(!prog_eqp, 3)) Then
'                MyIP = Val(Left(!prog_eqp, 3)) & "." & Val(Mid(!prog_eqp, 5, 3)) & "." & Val(Mid(!prog_eqp, 9, 3)) & "." & Val(Right(!prog_eqp, 3))
'            End If
'        End If
'    End If
'End With

Private Sub Form_Load()
    Dim Base As Database
    Dim Registro As Recordset
    Dim sPorts As String
    Dim nF As Integer
    
    If Dir(App.Path & "\dbase.mdb") = "" Then
        MsgBox "No Existe DBASE.MDB ...", vbCritical
        End
    End If
    Set Base = OpenDatabase(App.Path & "\dbase.mdb", False, False)
    Set Registro = Base.OpenRecordset("select * from equipos where (modelo <= 1 and not(isnull(prog_eqp)))")
    If Registro.EOF Then
        MsgBox "No existen Equipos con Direcciones IP válidas ó modelos EuroScale"
        Registro.Close
        Base.Close
        Set Registro = Nothing
        Set Base = Nothing
        End
    End If
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    If Dir(App.Path & "\nports.txt") <> "" Then
        nF = FreeFile()
        Open App.Path & "\sports.txt" For Input As #nF
        Line Input #nF, sPorts
        nPortS = Val(sPorts)
        Close #nF
    Else
        nPortS = 6173
    End If
    DataEqp.DatabaseName = App.Path & "\dbase.mdb"
    DataEqp.RecordSource = "select * from equipos"
    ChkSvr.Value = vbChecked
    CmdEnvTodas.Caption = "Enviar a Servidora."
    ChkSvr.Enabled = False
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub CmdImpTaras_Click()
'función específica para importar taras numéricas...
'fichero TARAS.DAT
'Tipo   Valor   CR+LF
'2Dig.  4Dig.   ---> Debe estar de acuerdo con step Balanza...
    
    Dim nFichTara As Integer
    Dim Modificacion As Boolean
    Dim Registro As Recordset
    Dim sTara As String
    Dim Base As Database
    Dim nT As Long
    Dim nV As Long
    Dim Bucle As Integer
    
    CmdImpTaras.Enabled = False
    If Dir(App.Path & "\taras.dat") = "" Then
        MsgBox "No existe Fichero Taras.dat"
        Exit Sub
    End If
    On Error GoTo errorlectura

    Set Base = OpenDatabase(App.Path & "\dbase.mdb", False, False)

    nFichTara = FreeFile()
    CadenadeLog "Importando Fichero : " & App.Path & "\taras.dat"

    Open (App.Path & "\taras.dat") For Input As nFichTara
    Do While Not EOF(nFichTara)
        Line Input #nFichTara, sTara
        If IsNumeric(Mid(sTara, 1, 2)) And IsNumeric(Mid(sTara, 3)) Then
            nT = Val(Mid(sTara, 1, 2))
            nV = Val(Mid(sTara, 3))
            If nT > 0 And nV < 10000 Then
                Set Registro = Base.OpenRecordset("select * from tipotara where tipo=" & CStr(nT))
                If Registro.EOF Then
                    Registro.AddNew
                    Registro.Fields("tipo") = nT
                    Registro.Fields("valor") = nV
                Else
                    Registro.Edit
                    Registro.Fields("tipo") = nT
                    Registro.Fields("valor") = nV
                End If
            
                Registro.Fields("tran_el") = " "
                Registro.Update
            
            End If
            
        
        End If
    Loop

errorlectura:
    CmdImpTaras.Enabled = True
    
    Close #nFichTara
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Fallo en la lectura del archivo taras.dat"
    Else
        On Error GoTo 0
        MsgBox "Fichero Importado."
    End If
End Sub


VERSION 5.00
Begin VB.Form Frmmain 
   BorderStyle     =   0  'None
   Caption         =   "HIPER & SECIN120.CFG & Free Text"
   ClientHeight    =   6975
   ClientLeft      =   1350
   ClientTop       =   1425
   ClientWidth     =   12390
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar Textos"
      Height          =   255
      Left            =   11040
      TabIndex        =   55
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar Copia"
      Height          =   255
      Left            =   8880
      TabIndex        =   54
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pegar"
      Height          =   255
      Left            =   11040
      TabIndex        =   53
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   255
      Left            =   8880
      TabIndex        =   52
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton CmdF 
      Caption         =   "Añadir/Modificar"
      Height          =   375
      Left            =   9600
      TabIndex        =   49
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   8
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   48
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   7
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   47
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   6
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   46
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   5
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   45
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   4
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   44
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   3
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   43
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   2
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   42
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   1
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   41
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox TxtF 
      Height          =   285
      Index           =   0
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   40
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox CmbSec 
      Height          =   315
      Left            =   7680
      TabIndex        =   30
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Guardar y Salir"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   8
      Left            =   4440
      TabIndex        =   28
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   27
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   6
      Left            =   4440
      TabIndex        =   26
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   25
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   24
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   23
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   22
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   21
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   20
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   240
      MaxLength       =   30
      TabIndex        =   19
      Top             =   6480
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   240
      MaxLength       =   30
      TabIndex        =   18
      Top             =   6120
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   240
      MaxLength       =   30
      TabIndex        =   17
      Top             =   5760
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   240
      MaxLength       =   30
      TabIndex        =   16
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   240
      MaxLength       =   30
      TabIndex        =   15
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   240
      MaxLength       =   30
      TabIndex        =   14
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   13
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   240
      MaxLength       =   30
      TabIndex        =   12
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   240
      MaxLength       =   30
      TabIndex        =   11
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   12240
      X2              =   7320
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      X1              =   7320
      X2              =   7320
      Y1              =   720
      Y2              =   6720
   End
   Begin VB.Label Label2 
      Caption         =   "Configuraciones Especiales."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   51
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Textos Libres Para ETQ./Sec."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   50
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label LblTf50 
      Caption         =   "Adicional 40:"
      Height          =   255
      Left            =   7800
      TabIndex        =   39
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label LblTf08 
      Caption         =   "Texto Libre 08:"
      Height          =   255
      Left            =   7680
      TabIndex        =   38
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label LblTf07 
      Caption         =   "Texto Libre 07:"
      Height          =   255
      Left            =   7680
      TabIndex        =   37
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label LblTf06 
      Caption         =   "Texto Libre 06:"
      Height          =   255
      Left            =   7680
      TabIndex        =   36
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LblTf05 
      Caption         =   "Texto Libre 05:"
      Height          =   255
      Left            =   7680
      TabIndex        =   35
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label LblTf04 
      Caption         =   "Texto Libre 04:"
      Height          =   255
      Left            =   7680
      TabIndex        =   34
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LblTf03 
      Caption         =   "Texto Libre 03:"
      Height          =   375
      Left            =   7680
      TabIndex        =   33
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LblTf02 
      Caption         =   "Texto Libre 02:"
      Height          =   255
      Left            =   7680
      TabIndex        =   32
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LblTf01 
      Caption         =   "Texto Libre 01:"
      Height          =   255
      Left            =   7680
      TabIndex        =   31
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label LblTipoEtq 
      Caption         =   "Secciones con Etiquetas 120mm.: (Marcar Secciones que utilizarán etiquetas de 120mm)"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   6615
   End
   Begin VB.Label LblPass 
      Caption         =   "Password:"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LblUser 
      Caption         =   "Usuario Cajas:"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LblFolder 
      Caption         =   "Carpeta Servidor Depósito Tiques:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label LblIpCaja 
      Caption         =   "IP Servidor Cajas:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LblHiper 
      Caption         =   "Identificación Hiper:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RECORDAR TRANSFORMAR TABLAS SI PRECISO Y CARGAR POR DEFECTO LOS TEXTOS DE CTF.FRE Y BORRAR CTF.FRE
Dim textC(10) As String
'Sección 26 a 3, Secciones > 9 y distintas a 26 eliminar
'Todos los plus a 0
Private Sub RevSec()
Dim MyBase As DAO.Database
Dim MyReg As DAO.Recordset
Dim MyRegSec As DAO.Recordset
Dim cm As String

    If Dir(App.Path & "\exporta", vbDirectory) <> "" Then
        On Error Resume Next
        If Dir(App.Path & "\exporta\*.*") <> "" Then
            Kill App.Path & "\exporta\*.*"
        End If
        On Error GoTo 0
    End If
    If Dir(App.Path & "\*.bck") <> "" Then
        Kill App.Path & "\*.bck"
    End If
    If Dir(App.Path & "\*.dat") <> "" Then
        Kill App.Path & "\*.dat"
    End If
    If Dir(App.Path & "\*.tmp") <> "" Then
        Kill App.Path & "\*.tmp"
    End If
    If Dir(App.Path & "\*.loc") <> "" Then
        Kill App.Path & "\*.loc"
    End If
    
    Set MyBase = OpenDatabase(App.Path & "\dbase.mdb")
    
        Set MyReg = MyBase.OpenRecordset("select * from seccion where secc_maqui<3")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
    
        Set MyReg = MyBase.OpenRecordset("select * from articulo where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Edit
                MyReg.Fields("tran_ep") = MyReg.Fields("secc_maqui")
                
                MyReg.Fields("secc_maqui") = 3
                MyReg.Fields("codi_ident") = 3
                MyReg.Fields("plu") = 0
                MyReg.Update
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from seccion where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from cabley where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from codbar where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from equipos where secc_maqui>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("secc_maqui") = 3
            MyReg.Fields("codi_ident") = 3
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from familia where secc_maqui>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("secc_maqui") = 3
            MyReg.Fields("codi_ident") = 3
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from fam_code where codi_ident>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("codi_ident") = 3
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from lintxt2040 where mostrador>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("mostrador") = 3
            MyReg.Fields("plu") = 0
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from publi where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from sub_code where codi_ident>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("codi_ident") = 3
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from subsec where codi_ident>9")
        Do While Not MyReg.EOF
            MyReg.Edit
            MyReg.Fields("codi_ident") = 3
            MyReg.Update
            MyReg.MoveNext
        Loop
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from teclas where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from textlibre where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing
        
        Set MyReg = MyBase.OpenRecordset("select * from vendedor where secc_maqui>9")
        If Not MyReg.EOF Then
            MyReg.MoveFirst
            Do While Not MyReg.EOF
                MyReg.Delete
                MyReg.MoveNext
            Loop
        End If
        MyReg.Close
        Set MyReg = Nothing

        MyBase.Execute "update articulo set plu=0 where codigo<>0"
        MyBase.Execute "delete from articulo where codigo=0"

End Sub
Public Sub Crea_TextLibre_1()
    '//////////////////////////////
    ' Textos Libres : similares
    ' a cabeceras y leyendas, excluisvas
    ' para Euroscale
    Dim Base As Database
    Dim Tabla As TableDef
    Dim bucle As Integer
    Dim i As Integer
    Dim reg As Recordset
    Set Base = OpenDatabase(App.Path & "\dbase.mdb")
    On Error Resume Next
    Set reg = Base.OpenRecordset("select texto49 from textlibre")
    If Err.Number <> 0 Then
        On Error GoTo 0
        Set Tabla = Base.TableDefs("textlibre")
        For bucle = 10 To 49
            Tabla.Fields.Append Tabla.CreateField("TEXTO" & Format(bucle, "00"), dbText, 40)
            Tabla.Fields("TEXTO" & Format(bucle, "00")).AllowZeroLength = True
        Next bucle
        Set Tabla = Base.TableDefs("fichavacuno")
        For bucle = 0 To 39
            Tabla.Fields.Append Tabla.CreateField("TEXTO" & Format(bucle, "00"), dbText, 40)
            Tabla.Fields("TEXTO" & Format(bucle, "00")).AllowZeroLength = True
        Next bucle
    End If
    Base.Close
End Sub

Private Sub Crea_1_7_0()
Dim db As Database
Dim tdf As TableDef
Dim rst As Recordset
Dim cnt As Long

    Set db = OpenDatabase(App.Path & "\dbase.mdb")
    On Error Resume Next
    Set rst = db.OpenRecordset("select * from gthora")
    If Err.Number <> 0 Then
        On Error GoTo 0
        Set tdf = db.CreateTableDef("gthora")
        With tdf
            .Fields.Append .CreateField("FECHA", dbDate)
            '.Fields.Append .CreateField("HORA", dbText)
            .Fields.Append .CreateField("HORA", dbDate)
            .Fields.Append .CreateField("SECCION", dbLong)
            .Fields.Append .CreateField("EQUIPO", dbLong)
            .Fields.Append .CreateField("VENDEDOR", dbLong)
            .Fields.Append .CreateField("TIQUET", dbLong)
            .Fields.Append .CreateField("OPERACIONES", dbLong)
            .Fields.Append .CreateField("OPERACIONES_CAN", dbLong)
            .Fields.Append .CreateField("OPERACIONES_NEG", dbLong)
            .Fields.Append .CreateField("IMPORTE", dbDouble, 9)
            .Fields.Append .CreateField("IMPORTE_CAN", dbDouble, 9)
            .Fields.Append .CreateField("IMPORTE_NEG", dbDouble, 9)
            .Fields.Append .CreateField("PESO", dbDouble)
            .Fields.Append .CreateField("PESO_CAN", dbDouble)
            .Fields.Append .CreateField("PESO_NEG", dbDouble)
            .Fields.Append .CreateField("UNIDADES", dbLong)
            .Fields.Append .CreateField("UNIDADES_CAN", dbLong)
            .Fields.Append .CreateField("UNIDADES_NEG", dbLong)
            db.TableDefs.Append tdf
        End With
    End If
    db.Close
End Sub

Private Sub RecargaFre()
Dim MyBase As DAO.Database
Dim MyReg As DAO.Recordset
Dim MyRegSec As DAO.Recordset
Dim nF As Integer
Dim sP, sT As String
Dim nC As Integer

    If Dir(App.Path & "\ctf.fre") <> "" Then
        On Error Resume Next
        Call Crea_1_7_0
        On Error GoTo 0
        
        nF = FreeFile()
        Open App.Path & "\ctf.fre" For Input As #nF
        Do While Not EOF(nF)
            Input #nF, sP, sT
            If Not IsNull(sT) Then
                If Val(sP) = 50 Then
                    sP = "9"
                    TxtF(Val(sP) - 1).Text = Trim(sT)
                Else
                    TxtF(Val(sP) - 1).Text = Trim(sT)
                End If
            End If
        Loop
        Close #nF
        
        Call Crea_TextLibre_1
        
        Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
        Set MyRegSec = MyBase.OpenRecordset("select * from seccion where secc_maqui<10 order by secc_maqui")
        MyRegSec.MoveFirst
        Do While Not MyRegSec.EOF
            Set MyReg = MyBase.OpenRecordset("select * from textlibre where secc_maqui=" & CStr(MyRegSec.Fields("secc_maqui")))
            If MyReg.EOF Then
                MyReg.AddNew
            Else
                MyReg.Edit
            End If
            MyReg.Fields("tienda") = 0
            MyReg.Fields("codi_ident") = MyRegSec.Fields("secc_maqui")
            MyReg.Fields("secc_maqui") = MyRegSec.Fields("secc_maqui")
            'NUMERO_EQP, 0
            MyReg.Fields("numero_eqp") = 0
            'SECC_EQP, False
            MyReg.Fields("secc_eqp") = False
            'texto00 a texto07
            MyReg.Fields("texto00") = Trim(TxtF(0).Text)
            MyReg.Fields("texto01") = Trim(TxtF(1).Text)
            MyReg.Fields("texto02") = Trim(TxtF(2).Text)
            MyReg.Fields("texto03") = Trim(TxtF(3).Text)
            MyReg.Fields("texto04") = Trim(TxtF(4).Text)
            MyReg.Fields("texto05") = Trim(TxtF(5).Text)
            MyReg.Fields("texto06") = Trim(TxtF(6).Text)
            MyReg.Fields("texto07") = Trim(TxtF(7).Text)
            MyReg.Fields("texto08") = " "
            MyReg.Fields("texto09") = Trim(Text1.Text)
            MyReg.Fields("texto49") = Trim(TxtF(8).Text)
            For nC = 0 To 9
                MyReg.Fields("TLETRA" & Format(nC, "00")) = 0
            Next nC
            MyReg.Fields("TRAN_TXT") = " "
            MyReg.Fields("TRAN_TXTSC10") = " "
            'BORRADO
            MyReg.Fields("BORRADO") = False
            For nC = 10 To 48
                MyReg.Fields("texto" & Format(nC, "00")) = " "
            Next nC
            MyReg.Update
    
            MyReg.Close
            MyRegSec.MoveNext
        Loop
        MyRegSec.Close
        MyBase.Close
    End If
End Sub

Private Sub Recarga()
Dim MyBase As DAO.Database
Dim MyReg As DAO.Recordset
Dim MyRegF As DAO.Recordset
'Dim nC As Integer
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    
    Set MyReg = MyBase.OpenRecordset("select * from seccion where secc_maqui<10 order by secc_maqui")
    If MyReg.EOF Then
        MsgBox "No existen Secciones...", vbCritical
        End
    End If
    MyReg.MoveFirst
    'nC = 0
    Do While Not MyReg.EOF
        Set MyRegF = MyBase.OpenRecordset("select * from textlibre where secc_maqui=" & CStr(MyReg.Fields("secc_maqui")))
        If Not MyRegF.EOF Then
            'On Error Resume Next
            MyFreeT(MyReg.Fields("secc_maqui")).text01 = MyRegF.Fields("texto00")
            'On Error GoTo 0
            'On Error Resume Next
            MyFreeT(MyReg.Fields("secc_maqui")).text02 = MyRegF.Fields("texto01")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text03 = MyRegF.Fields("texto02")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text04 = MyRegF.Fields("texto03")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text05 = MyRegF.Fields("texto04")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text06 = MyRegF.Fields("texto05")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text07 = MyRegF.Fields("texto06")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).text08 = MyRegF.Fields("texto07")
            'On Error GoTo 0
            'On Error Resume Next
            
            MyFreeT(MyReg.Fields("secc_maqui")).adicional40 = MyRegF.Fields("texto49")
            'On Error GoTo 0
        Else
            MyFreeT(MyReg.Fields("secc_maqui")).text01 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text02 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text03 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text04 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text05 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text06 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text07 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).text08 = ""
            MyFreeT(MyReg.Fields("secc_maqui")).adicional40 = ""
        End If
        MyRegF.Close
        MyReg.MoveNext
        'nC = nC + 1
    Loop
    MyReg.Close
    MyBase.Close

End Sub

Private Sub CmbSec_Click()
    TxtF(0).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text01
    TxtF(1).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text02
    TxtF(2).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text03
    TxtF(3).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text04
    TxtF(4).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text05
    TxtF(5).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text06
    TxtF(6).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text07
    TxtF(7).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).text08
    TxtF(8).Text = MyFreeT(Val(Mid(CmbSec.Text, 1, 3))).adicional40
    If Trim(TxtF(8).Text) = "" Then
        TxtF(8).Text = "F.Caducidad"
    End If
    
End Sub

Private Sub CmdF_Click()
Dim MyBase As DAO.Database
Dim MyReg As DAO.Recordset
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyReg = MyBase.OpenRecordset("select * from textlibre where secc_maqui=" & CStr(Val(Left(CmbSec.Text, 3))))
    If MyReg.EOF Then
        MyReg.AddNew
    Else
        MyReg.Edit
    End If
    MyReg.Fields("tienda") = 0
    MyReg.Fields("codi_ident") = Val(Left(CmbSec.Text, 3))
    MyReg.Fields("secc_maqui") = Val(Left(CmbSec.Text, 3))
    'NUMERO_EQP, 0
    MyReg.Fields("numero_eqp") = 0
    'SECC_EQP, False
    MyReg.Fields("secc_eqp") = False
    'texto00 a texto07
    MyReg.Fields("texto00") = Trim(TxtF(0).Text)
    MyReg.Fields("texto01") = Trim(TxtF(1).Text)
    MyReg.Fields("texto02") = Trim(TxtF(2).Text)
    MyReg.Fields("texto03") = Trim(TxtF(3).Text)
    MyReg.Fields("texto04") = Trim(TxtF(4).Text)
    MyReg.Fields("texto05") = Trim(TxtF(5).Text)
    MyReg.Fields("texto06") = Trim(TxtF(6).Text)
    MyReg.Fields("texto07") = Trim(TxtF(7).Text)
    MyReg.Fields("texto08") = ""
    MyReg.Fields("texto09") = Trim(Text1.Text)
    MyReg.Fields("texto49") = Trim(TxtF(8).Text)
    MyReg.Update
    MsgBox "Textos añadidos o modificados.", vbExclamation
    MyReg.Close
    MyBase.Close
    
    Call Recarga
End Sub

Private Sub CmdSave_Click()
Dim nF As Integer
Dim nC As Integer
    CmdSave.Enabled = False
    nF = FreeFile()
    Open App.Path & "\hiper" For Output As #nF
    Print #nF, Trim(Text1.Text)
    Print #nF, Trim(Text2.Text)
    Print #nF, Trim(Text3.Text)
    Print #nF, Trim(Text4.Text)
    Print #nF, Trim(Text5.Text)
    Close #nF
    nF = FreeFile()
    Open App.Path & "\secin120.cfg" For Output As #nF
    For nC = 0 To 8
        If Check1(nC).Value = vbChecked Then
            Print #nF, CStr(Left(Text6(nC).Text, 2))
        End If
    Next nC
    Close #nF
    Call Corregir_Nulos
    End
End Sub

Private Sub Command1_Click()
    textC(0) = TxtF(0).Text
    textC(1) = TxtF(1).Text
    textC(2) = TxtF(2).Text
    textC(3) = TxtF(3).Text
    textC(4) = TxtF(4).Text
    textC(5) = TxtF(5).Text
    textC(6) = TxtF(6).Text
    textC(7) = TxtF(7).Text
    textC(8) = TxtF(8).Text

End Sub

Private Sub Command2_Click()
    TxtF(0).Text = textC(0)
    TxtF(1).Text = textC(1)
    TxtF(2).Text = textC(2)
    TxtF(3).Text = textC(3)
    TxtF(4).Text = textC(4)
    TxtF(5).Text = textC(5)
    TxtF(6).Text = textC(6)
    TxtF(7).Text = textC(7)
    TxtF(8).Text = textC(8)
End Sub

Private Sub Command3_Click()
    textC(0) = ""
    textC(1) = ""
    textC(2) = ""
    textC(3) = ""
    textC(4) = ""
    textC(5) = ""
    textC(6) = ""
    textC(7) = ""
    textC(8) = ""
    textC(9) = ""
End Sub

Private Sub Command4_Click()
    TxtF(0).Text = ""
    TxtF(1).Text = ""
    TxtF(2).Text = ""
    TxtF(3).Text = ""
    TxtF(4).Text = ""
    TxtF(5).Text = ""
    TxtF(6).Text = ""
    TxtF(7).Text = ""
    TxtF(8).Text = ""
End Sub

Private Sub Form_Load()
Dim nF As Integer
Dim nC As Integer
Dim nPos As Integer
Dim MyBase As DAO.Database
Dim MyReg As DAO.Recordset
Dim MyRegF As DAO.Recordset
Dim sH As String
Dim sI As String
Dim sF As String
Dim sU As String
Dim sP As String
Dim sS As String

'Hiper
sH = "0028"
sI = "206.0.52.3"
sF = "/home/cajas/tickbal/"
sU = "cajas"
sP = "cajas"

If Dir(App.Path & "\dbase.mdb") <> "" Then
    Call RevSec
    If Dir(App.Path & "\hiper") <> "" Then
        nF = FreeFile()
        Open App.Path & "\hiper" For Input As #nF
        nC = 0
        Do While Not EOF(nF)
            Select Case nC
            Case 0
                Line Input #nF, sS
                Text1.Text = Trim(sS)
                If Len(Text1.Text) < 4 Then
                    Text1.Text = Mid("0000", 1, 4 - Len(Text1.Text)) & Text1.Text
                End If
            Case 1
                Line Input #nF, sS
                Text2.Text = Trim(sS)
            Case 2
                Line Input #nF, sS
                Text3.Text = Trim(sS)
            Case 3
                Line Input #nF, sS
                Text4.Text = Trim(sS)
            Case 4
                Line Input #nF, sS
                Text5.Text = Trim(sS)
            End Select
            nC = nC + 1
        Loop
        Close #nF
        If nC < 4 Then
            
            Text2.Text = sI
            Text3.Text = sF
            Text4.Text = sU
            Text5.Text = sP
            If Dir(App.Path & "\caja.bat") <> "" Then
                nF = FreeFile()
                Open App.Path & "\caja.bat" For Input As #nF
                nC = 0
                Do While Not EOF(nF)
                    Line Input #nF, sS
                    nC = nC + 1
                    If nC = 2 Then
                        nPos = InStr(1, sS, "cajas.dat")
                        If nPos <> 0 Then
                            sI = Trim(Mid(sS, nPos + 9))
                            Text2.Text = sI
                        End If
                    End If
                Loop
                Close #nF
            End If
            
        End If
    Else
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
    End If
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "NO ES POSIBLE ABRIR DBASE.MDB...", vbCritical
        End
    End If
    On Error GoTo 0
    Set MyReg = MyBase.OpenRecordset("select * from seccion where secc_maqui<10 order by secc_maqui")
    If MyReg.EOF Then
        MsgBox "No existen Secciones...", vbCritical
        End
    End If
    MyReg.MoveFirst
    nC = 0
    Do While Not MyReg.EOF
        Text6(nC).Text = Format(CStr(MyReg.Fields("secc_maqui")), "00") & " -" & Trim(MyReg.Fields("descripcio"))
        CmbSec.AddItem Text6(nC).Text, nC
        MyReg.MoveNext
        nC = nC + 1
    Loop
    MyReg.Close
    MyBase.Close
    CmbSec.ListIndex = 0
    Call RecargaFre
    Call Recarga
    Call CmbSec_Click
    
    If Dir(App.Path & "\secin120.cfg") <> "" Then
        nF = FreeFile()
        Open App.Path & "\secin120.cfg" For Input As #nF
        Do While Not EOF(nF)
            Line Input #nF, sS
            sS = Format(Trim(sS), "00")
            For nC = 0 To 8
                If Trim(Text6(nC).Text) <> "" Then
                    If Val(Trim(Left(Text6(nC).Text, 3))) = Val(sS) Then Check1(nC).Value = vbChecked
                End If
            Next nC
        Loop
        Close #nF
    End If
    For nC = 0 To 8
        If Trim(Text6(nC).Text) = "" Then Check1(nC).Enabled = False
    Next nC
Else
    End
End If
End Sub

Private Sub Text1_LostFocus()
    If Len(Text1.Text) <> 4 Then
        Text1.Text = Trim(Text1.Text)
        Text1.Text = Mid("0000", 1, 4 - Len(Text1.Text)) & Trim(Text1.Text)
    End If
End Sub

Public Sub Corregir_Nulos()
'/////////////////////////////////////
'// esta función es llamada al inicio
'// del programa para eliminar los posibles
'// valores nulos en campos 'dbText' de las tablas, ya que
'// algunos usuarios pican artículos
'// directamente en la base, dejando
'// los campos de texto como nulos.
'// Se podría implementar también este
'// proceso para valores numéricos, pero
'// hay que tener mucho cuidado con los
'// valores por defecto que se asignan.
'//////////////////////////////////////
Dim Base As Database
Dim rst As Recordset
Dim cntTab As Long
Dim cntReg As Long
Dim cntCam As Long
Dim StrSQL As String
    
    On Error GoTo ret1
    Set Base = OpenDatabase(App.Path & "\dbase.mdb")
    Base.Execute "update articulo set etq=255 where etq<>255"
    Base.Execute "update cabecera set factura=' ' where isnull(factura)"
    Base.Execute "update tickets set factura=' ' where isnull(factura)"
    Base.Execute "update cabecerasc10 set factura=' ' where isnull(factura)"
    Base.Execute "update ticketssc10 set factura=' ' where isnull(factura)"
    Base.Execute "update cabeceratq set factura=' ' where isnull(factura)"
    Base.Execute "update ticketstq set factura=' ' where isnull(factura)"
    
    For cntTab = 0 To Base.TableDefs.Count - 1
        On Error GoTo ret1
        If Left(UCase(Base.TableDefs(cntTab).Name), 4) <> "MSYS" Then
            StrSQL = "SELECT * FROM " & Base.TableDefs(cntTab).Name
            '2.0.65
            'If UCase(Mid(Base.TableDefs(cntTab).Name, 1, 8)) <> "CABECERA" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 7)) <> "TICKETS" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 3)) <> "LOG" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 6)) <> "TECLAS" Then
            '   'UCase(Mid(Base.TableDefs(cntTab).Name, 1, 8)) <> "ARTICULO" And
            If UCase(Mid(Base.TableDefs(cntTab).Name, 1, 3)) <> "LOG" And _
               UCase(Mid(Base.TableDefs(cntTab).Name, 1, 6)) <> "TECLAS" Then
            On Error Resume Next
            Set rst = Base.OpenRecordset(StrSQL)
            If Err.Number <> 0 Then
                On Error GoTo 0
                GoTo SigueTabla
            End If
            If Not rst.BOF Then
                rst.MoveFirst
                Do Until rst.EOF
                    For cntCam = 0 To rst.Fields.Count - 1
                        'If LCase(rst.Fields(cntCam).Name) = "des_plu2" Then
                        '    cntCam = cntCam
                        'End If
                        
                        If rst.Fields(cntCam).Type = dbText Then
                            If IsNull(rst.Fields(cntCam)) Then
                                rst.Edit
                                On Error Resume Next
                                rst.Fields(cntCam) = ""
                                If Err.Number <> 0 Then
                                    On Error GoTo 0
                                    On Error GoTo ret1
                                    If IsNull(rst.Fields(cntCam)) And LCase(rst.Fields(cntCam).Name) = "factura" Then
                                         rst.Fields(cntCam) = " "
                                    End If
                                End If
                                On Error GoTo ret1
                                rst.Update
                            End If
                        Else
                            If rst.Fields(cntCam).Type = dbNumeric Or rst.Fields(cntCam).Type = 4 Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).Type = dbBoolean Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    rst.Edit
                                    rst.Fields(cntCam) = False
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).Type = dbDouble Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).Type = dbBigInt Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            
                        
                        End If
                    Next cntCam
                    'Do_Events
                    rst.MoveNext
                Loop
            End If
            rst.Close
            Set rst = Nothing
            End If
        End If
SigueTabla:
    Next cntTab
    
ret1:
On Error GoTo 0
On Error GoTo Ret
    Base.Execute "update cabecera set imp0=0 where base0=0 or isnull(base0)"
    Base.Execute "update cabecera set imp1=0 where base1=0 or isnull(base1)"
    Base.Execute "update cabecera set imp2=0 where base2=0 or isnull(base2)"
    Base.Execute "update cabecera set imp3=0 where base3=0 or isnull(base3)"
    Base.Execute "update cabecera set imp4=0 where base4=0 or isnull(base4)"
    
    Base.Execute "update cabecerasc10 set imp0=0 where base0=0 or isnull(base0)"
    Base.Execute "update cabecerasc10 set imp1=0 where base1=0 or isnull(base1)"
    Base.Execute "update cabecerasc10 set imp2=0 where base2=0 or isnull(base2)"
    Base.Execute "update cabecerasc10 set imp3=0 where base3=0 or isnull(base3)"
    Base.Execute "update cabecerasc10 set imp4=0 where base4=0 or isnull(base4)"
    
    Base.Execute "update cabeceratq set imp0=0 where base0=0 or isnull(base0)"
    Base.Execute "update cabeceratq set imp1=0 where base1=0 or isnull(base1)"
    Base.Execute "update cabeceratq set imp2=0 where base2=0 or isnull(base2)"
    Base.Execute "update cabeceratq set imp3=0 where base3=0 or isnull(base3)"
    Base.Execute "update cabeceratq set imp4=0 where base4=0 or isnull(base4)"
    
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    Corregir_Articulos_Nulos
    
    Corregir_Trazabilidad_Nulos
    
    On Error GoTo 0
Exit Sub
Ret:
    
    '1.7.2 If TypeName(Base) <> "Nothing" Then Base.Close
    rst.Close
    Set rst = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    On Error GoTo 0
End Sub
Private Sub Corregir_Articulos_Nulos()
    Dim Base As Database
    Dim bucle As Integer
    Dim Registro As Recordset
    Dim Campos(22) As String
    
    Campos(0) = "codi_ident"
    Campos(1) = "plu"
    Campos(2) = "codigo"
    Campos(3) = "secc_maqui"
    Campos(4) = "codi_sub"
    Campos(5) = "codi_fam"
    Campos(6) = "tipo_iva"
    Campos(7) = "caducidad"
    Campos(8) = "grupo_conserv"
    Campos(9) = "tara"
    Campos(10) = "pref"
    Campos(11) = "fcb"
    Campos(12) = "ning"
    Campos(13) = "precio"
    Campos(14) = "euros"
    Campos(15) = "usatramos"
    Campos(16) = "borrado"
    Campos(17) = "balenv"
    Campos(18) = "prc100g"
    Campos(19) = "etq"
    Campos(20) = "tip_let0"
    Campos(21) = "poid"
    Set Base = OpenDatabase(App.Path & "\dbase.mdb")
    For bucle = 0 To 21
        Set Registro = Base.OpenRecordset("select * from articulo where isnull(" & Campos(bucle) & ")")
        With Registro
            If Not .EOF Then
                .MoveFirst
                If bucle < 4 Then
                    Do Until .EOF
                        If Not .EOF Then .Delete
                        If Not .EOF Then .MoveNext
                    Loop
                Else
                    Do Until .EOF
                        If Not .EOF Then .Edit
                        Select Case bucle
                            Case 13 To 14
                                .Fields(Campos(bucle)) = 1
                            Case 15 To 18
                                .Fields(Campos(bucle)) = False
                            Case 19
                                .Fields(Campos(bucle)) = 255 ' ETQ
                            Case 21
                                .Fields(Campos(bucle)) = 0
                            Case Else
                                .Fields(Campos(bucle)) = 0
                        End Select
                        .Update
                        If Not .EOF Then .MoveNext
                    Loop
                End If
            End If
        End With
    Next bucle
    
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    
End Sub

Public Sub Corregir_Trazabilidad_Nulos(Optional ByVal nF As Integer)
    Dim Base As Database
    Dim bucle As Integer
    Dim Registro As Recordset
    Dim rPa As Recordset
    Dim Campos(22) As String
    Dim s As String
    
    
    Campos(0) = "codnacimiento"
    Campos(1) = "codcrianza"
    Campos(2) = "codsacrificio"
    Campos(3) = "coddespiece"
    Campos(4) = "rsisacrificio"
    Campos(5) = "rsidespiece"
    Campos(6) = "fecha_nac"
    Campos(7) = "fecha_sac"
    Campos(8) = "fecha_des"
    Set Base = OpenDatabase(App.Path & "\dbase.mdb")

    If IsNull(nF) Or nF = 0 Then
        Set Registro = Base.OpenRecordset("select * from fichavacuno where texto00='' or trim(texto00)=''")
    Else
        Set Registro = Base.OpenRecordset("select * from fichavacuno where codigo='" & CStr(nF) & "'")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF()
                .Edit
                Set rPa = Base.OpenRecordset("select * from paises where codigo=" & CStr(.Fields(Campos(0))))
                If Not rPa.EOF() Then
                    s = "Nacimiento:" & Trim(rPa.Fields("nombre"))
                Else
                    s = ""
                End If
                rPa.Close
                .Fields("texto00") = s
                Set rPa = Base.OpenRecordset("select * from paises where codigo=" & CStr(.Fields(Campos(1))))
                If Not rPa.EOF() Then
                    s = "Crianza:" & Trim(rPa.Fields("nombre"))
                Else
                    s = ""
                End If
                rPa.Close
                .Fields("texto01") = s
                Set rPa = Base.OpenRecordset("select * from paises where codigo=" & CStr(.Fields(Campos(2))))
                If Not rPa.EOF() Then
                    s = "Sacrificio:" & Trim(rPa.Fields("nombre"))
                Else
                    s = ""
                End If
                rPa.Close
                .Fields("texto02") = s
                Set rPa = Base.OpenRecordset("select * from paises where codigo=" & CStr(.Fields(Campos(3))))
                If Not rPa.EOF() Then
                    s = "Despiece:" & Trim(rPa.Fields("nombre"))
                Else
                    s = ""
                End If
                rPa.Close
                .Fields("texto03") = s
                .Fields("texto04") = "RSI Sacrificio:" & Trim(.Fields(Campos(4)))
                .Fields("texto05") = "RSI Despiece:" & Trim(.Fields(Campos(5)))
                .Fields("texto06") = "F.Nacimiento:" & Trim(.Fields(Campos(6)))
                .Fields("texto07") = "F.Sacrificio:" & Trim(.Fields(Campos(7)))
                .Fields("texto08") = "F.Despiece:" & Trim(.Fields(Campos(8)))
                    
                .Update
                
                .MoveNext
            Loop
        End If
    End With

    
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    
End Sub


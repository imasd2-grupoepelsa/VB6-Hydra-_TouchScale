VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Procesos by ThePirat 2004-2007"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDrag 
      Height          =   660
      Left            =   7800
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   600
      ScaleWidth      =   1365
      TabIndex        =   18
      ToolTipText     =   "Haga click aquí y arrastre el mouse hasta la ventana que desea ubicar en la lista"
      Top             =   0
      Width           =   1425
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      ToolTipText     =   "Cierra el programa"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventana:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   7800
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
      Begin VB.CommandButton cmdFree 
         Caption         =   "Liberar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Deshabilita la entrada de teclado y mouse a la ventana"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdParent 
         Caption         =   "Cambiar Padre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Deshabilita la entrada de teclado y mouse a la ventana"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinClose 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Cierra la ventana"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinShow 
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Hace visible la ventana"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinHide 
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Hace invisible la ventana"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinDisable 
         Caption         =   "Deshabilitar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Deshabilita la entrada de teclado y mouse a la ventana"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinEnable 
         Caption         =   "Habilitar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Habilita la entrada de teclado y mouse a la ventana"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinRestore 
         Caption         =   "Restaurar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Restaura la ventana"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinMin 
         Caption         =   "Minimizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Minimiza la ventana"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdWinMax 
         Caption         =   "Maximizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Maximiza la ventana"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdFocus 
         Caption         =   "Enfocar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Devuelve el foco a la ventana"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdTopMost 
         Caption         =   "TopMost"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Coloca a la ventana seleccionada por encima de todas"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdNrm 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Vuelve la ventana a la normalidad"
         Top             =   735
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      Begin VB.CommandButton cmdKill 
         Caption         =   "MATAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Mata al proceso seleccionado"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   "MÁS INFO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Brinda información detallada del proceso"
         Top             =   735
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      ToolTipText     =   "Actualiza la lista"
      Top             =   8400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   16113
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Por Federico Colombo. The-Pirat 2005. thepirat000@hotmail.com
Option Explicit

Const sCaptionVentana = "Procesos"

Dim m_bSeleccionandoVentana As Boolean
Dim m_lVentanaHija As Long
Dim m_lVentanaPadre As Long

Dim Curs As POINTAPI
Dim CurrWin As Long, RowNegritaActual(0 To 1) As Integer

'Salir
Private Sub cmdClose_Click()
End
End Sub

'Enfocar ventana seleccionada
Private Sub cmdFocus_Click()
Dim handleW As Long
handleW = DameHwnd
SetWindowPos handleW, -1, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOCOPYBITS
SetWindowPos handleW, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOCOPYBITS
SetForegroundWindow handleW
End Sub

'Matar proceso seleccionado
Private Sub cmdKill_Click()
If MsgBox("Matar: " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), vbYesNo, "Confirmación") = vbNo Then Exit Sub
Call Matar(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
Call MostrarWindows
End Sub

'Mostrar información del proceso seleccionado
Private Sub cmdMore_Click()
MsgBox MasInfo(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)), vbInformation, "Detalles del proceso: " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End Sub

'Cambiar ventana seleccionada a normal
Private Sub cmdNrm_Click()
SetWindowPos DameHwnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOCOPYBITS
End Sub

'Cambiar padre de ventana seleccionada
Private Sub cmdParent_Click()
fbSeleccionVentana Not m_bSeleccionandoVentana
End Sub

'Liberar ventana seleccionada (quitar padre)
Private Sub cmdFree_Click()
fbModificarPadreVentana m_lVentanaHija, 0
fbSeleccionVentana False
End Sub

'Función que asigna nuevo padre a una ventana
Private Function fbSeleccionVentana(bSeleccion As Boolean) As Boolean
m_bSeleccionandoVentana = bSeleccion
If m_bSeleccionandoVentana Then
    cmdParent.Caption = "Asignar Padre"
    cmdFree.Enabled = True
    Me.Caption = sCaptionVentana & " - Seleccione la ventana padre para --> " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
    m_lVentanaHija = CLng(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
    fbSeleccionVentana = True
Else
    If MSFlexGrid1.Col < 2 Then
        MsgBox "Debe seleccionar una ventana, no un proceso. O bien seleccionar la opción de Liberar", vbInformation
        Exit Function
    End If
    cmdParent.Caption = "Cambiar Padre"
    cmdFree.Enabled = False
    Me.Caption = sCaptionVentana
    'Asignar el padre
    m_lVentanaPadre = CLng(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
    fbSeleccionVentana = fbModificarPadreVentana(m_lVentanaHija, m_lVentanaPadre)
End If

cmdWinClose.Enabled = Not m_bSeleccionandoVentana
cmdWinDisable.Enabled = Not m_bSeleccionandoVentana
cmdWinEnable.Enabled = Not m_bSeleccionandoVentana
cmdWinHide.Enabled = Not m_bSeleccionandoVentana
cmdWinMax.Enabled = Not m_bSeleccionandoVentana
cmdWinMin.Enabled = Not m_bSeleccionandoVentana
cmdWinRestore.Enabled = Not m_bSeleccionandoVentana
cmdWinShow.Enabled = Not m_bSeleccionandoVentana
cmdTopMost.Enabled = Not m_bSeleccionandoVentana
cmdNrm.Enabled = Not m_bSeleccionandoVentana
cmdFocus.Enabled = Not m_bSeleccionandoVentana
cmdRefresh.Enabled = Not m_bSeleccionandoVentana
End Function

'Volver a cargar el grid
Private Sub cmdRefresh_Click()
MostrarWindows
End Sub

'Cambiar ventana seleccionada a TopMost
Private Sub cmdTopMost_Click()
SetWindowPos DameHwnd, -1, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOCOPYBITS
End Sub

'Cerrar ventana seleccionada
Private Sub cmdWinClose_Click()
WindowHandle DameHwnd, eWM_CLOSE
Call MostrarWindows
End Sub

'Deshabilitar ventana seleccionada
Private Sub cmdWinDisable_Click()
windowEnable DameHwnd, False
Call MostrarWindows
End Sub

'Habilitar ventana seleccionada
Private Sub cmdWinEnable_Click()
windowEnable DameHwnd, True
Call MostrarWindows
End Sub

'Ocultar ventana seleccionada
Private Sub cmdWinHide_Click()
WindowHandle DameHwnd, eSW_HIDE
Call MostrarWindows
End Sub

'Cambiar ventana seleccionada a Maximizado
Private Sub cmdWinMax_Click()
WindowHandle DameHwnd, eSW_MAXIMIZE
End Sub

Private Function DameHwnd() As Long
DameHwnd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
End Function

'Cambiar ventana seleccionada a Minimizado
Private Sub cmdWinMin_Click()
WindowHandle DameHwnd, eSW_MINIMIZE
End Sub

'Cambiar ventana seleccionada a Restaurar
Private Sub cmdWinRestore_Click()
WindowHandle DameHwnd, eSW_RESTORE
End Sub

'Mostrar ventana seleccionada
Private Sub cmdWinShow_Click()
WindowHandle DameHwnd, eSW_SHOW
Call MostrarWindows
End Sub

Private Sub Form_Load()
Me.Caption = "Procesos by The-Pirat 2005 v" & App.Major & "." & App.Minor
MostrarWindows
End Sub

'Carga y formatea el grid
Private Function MostrarWindows()
Screen.MousePointer = vbHourglass
MSFlexGrid1.Visible = False
ActualizarObjetosWMI
With MSFlexGrid1
  .Clear
  .Rows = 1
  .Cols = 5
  .ColAlignment(1) = 1
  .ColAlignment(3) = 1
  .ColWidth(0) = 800
  .ColWidth(1) = 1500
  .ColWidth(2) = 700
  .ColWidth(3) = 3500
  .ColWidth(4) = 700
  .TextMatrix(0, 0) = "Proceso"
  .TextMatrix(0, 1) = "Nombre"
  .TextMatrix(0, 2) = "Ventana"
  .TextMatrix(0, 3) = "Título"
  .TextMatrix(0, 4) = "Estado"
  EnumWindows AddressOf EnumWindowsProc, ByVal 0&
  .Col = 2
  .ColSel = 0
  .Sort = 1
  .Row = 1
  .RowSel = 1
  .MergeCells = flexMergeFree
  .MergeCol(0) = True
  .MergeCol(1) = True
  .Col = 0
  .ColSel = 0
  Call MSFlexGrid1_SelChange
End With
Screen.MousePointer = vbDefault
MSFlexGrid1.Visible = True
End Function

Private Sub Form_Resize()
On Error Resume Next
Me.Width = 9435
MSFlexGrid1.Height = Me.Height - 400
End Sub

Private Sub MSFlexGrid1_DblClick()
If MSFlexGrid1.Col < 2 Then MsgBox colEjecutable.Item("K" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)), vbInformation, "Ruta del ejecutable"
End Sub

Private Sub MSFlexGrid1_SelChange()
Dim tmp As Boolean

If m_bSeleccionandoVentana Then Exit Sub

tmp = IIf(MSFlexGrid1.ColSel = 0 Or MSFlexGrid1.ColSel = 1, False, True)

If MSFlexGrid1.ColSel = 4 Then
    MSFlexGrid1.ToolTipText = "I=Invisible, V=Visible, E=Enabled, D=Disabled"
Else
    MSFlexGrid1.ToolTipText = ""
End If
cmdTopMost.Enabled = tmp
cmdNrm.Enabled = tmp
cmdFocus.Enabled = tmp
cmdWinMax.Enabled = tmp
cmdWinMin.Enabled = tmp
cmdWinRestore.Enabled = tmp
cmdWinEnable.Enabled = tmp
cmdWinDisable.Enabled = tmp
cmdWinClose.Enabled = tmp
cmdWinHide.Enabled = tmp
cmdWinShow.Enabled = tmp
cmdParent.Enabled = tmp
cmdKill.Enabled = Not tmp
cmdMore.Enabled = Not tmp

End Sub

'Click en el recuadro para seleccionar ventanas
Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbCrosshair
End Sub

'Al soltar luego de arrastrar desde el recuadro para seleccionar ventanas
'Busca y selecciona la ventana en el grid.
Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetCursorPos Curs
CurrWin = WindowFromPoint(Curs.X, Curs.Y)
Screen.MousePointer = vbDefault
Dim i As Integer, j As Byte, PID As Long, YaPid As Boolean

GetWindowThreadProcessId CurrWin, PID

With MSFlexGrid1
    If RowNegritaActual(0) <> 0 Then
          .Row = RowNegritaActual(0)
            For j = 0 To 1
                .Col = j
                .CellFontBold = False
            Next j
    End If
    
    If RowNegritaActual(1) <> 0 Then
          .Row = RowNegritaActual(1)
            For j = 2 To 3
                .Col = j
                .CellFontBold = False
            Next j
    End If
    
    YaPid = False
    For i = 1 To .Rows - 1
        
        If .TextMatrix(i, 0) = PID And Not YaPid Then
            .Row = i
            For j = 0 To 1
                .Col = j
                .CellFontBold = True
            Next j
            RowNegritaActual(0) = i
            YaPid = True
            
            .RowSel = i
            .Row = i
            .Col = 3
            .ColSel = 3
            .SetFocus
            .TopRow = i
            Call MSFlexGrid1_SelChange
            
        End If
        
        If .TextMatrix(i, 2) = CurrWin Then
            .Row = i
            For j = 2 To 3
                .Col = j
                .CellFontBold = True
            Next j
            RowNegritaActual(1) = i
        End If
    
    Next i
    .Refresh
End With
End Sub




Attribute VB_Name = "Alcampo_Especial"
'///////////////////////////////////
' estas funciones se crearon en principio
' para Alcampo (Ver. 4.5.18) pero
' han quedado en el estandar
'////////////////////////////////////
Option Explicit
Public Sub MS_EnviaBaseRemota(Optional Bloqueo As Boolean)
    '//////////////////////////
    ' replica/actualiza la base de datos
    ' con el otro equipo en sistemas
    ' Master/Slave
    '///////////////////////////
    Dim BufL As Boolean
    Dim MiProceso As Long
    Dim MiResultado As Long
    Dim Bucle As Integer
    If MS_Lock Then
        Exit Sub
    End If
    MS_Lock = True
    BufL = lCogeTiquet
    '/////////////////////////
    ' Se evita que se lance
    ' la recepción de tiquets contínua
    lCogeTiquet = False
    '////////////////////////////////
    ' El proceso se realiza realmente
    ' a través del ejecutable
    ' mastercopy.exe, ya que de hacerse
    ' la llamada de replicación desde
    ' el propio programa, este queda sin
    ' atender al socket de sincronización
    ' Master/Slave, con lo cual se puede
    ' propiciar el lanzamiento del esclavo
    ' de modo incorrecto si el proceso tarda
    ' demasiado
    If Dir(MiruTa & "\mastercopy.exe") <> "" Then
        'If Bloqueo = False Then
        frmEpelsa.LblSC10.Caption = ""
        frmEpelsa.LblSC10.ForeColor = vbGreen
        On Error Resume Next
        Load Form2
        Form2.Label1.Caption = "Transfiriendo Datos al otro equipo..."
        Form2.Show
        Do_Events
        ChDir MS_RPath
        On Error GoTo 0
        ChDir App.Path
        Do_Events
        MiResultado = STILL_ACTIVE
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiruTa & "\mastercopy.exe " & MS_RPath, vbNormalFocus))
        If MiProceso <> 0 Then
            Do While MiResultado = STILL_ACTIVE
                Do_Events
                GetExitCodeProcess MiProceso, MiResultado
            Loop
        End If
    Else
        MsgBox "MASTERCOPY.EXE NOT FOUND. CAN'T CONTINUE", vbCritical
        salir_programa
    End If
    MS_Lock = False
    lCogeTiquet = BufL
    Unload Form2
    On Error GoTo 0
End Sub

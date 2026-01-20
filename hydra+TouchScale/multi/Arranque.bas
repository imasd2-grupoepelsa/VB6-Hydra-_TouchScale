Attribute VB_Name = "Arranque"
Option Explicit
Public ID As Integer
Public HaySeguridad As Boolean
Public Backup_Path As String
Public Exp_Fecha As Integer
Public Exp_Hora As Integer
Public Exp_Absoluto As Boolean
Public Exp_Diario As Boolean
Public Exp_Positivo As String
Public Exp_Negativo As String
Public Const MiVer As String = "Release 2.0.0"
Sub Main()
    If Dir(App.Path & "\hydra.exe") = "" And Dir(App.Path & "\hydratouch.exe") = "" Then
        MsgBox "THIS PROGRAM MUST BE INSTALLED IN THE SAME FOLDER AS HYDRATOUCH", vbCritical
        End
    End If
    Do_Events
    If App.PrevInstance Then End
    'If Trim(Command) <> "" Then
    '    Shell (App.Path & "\hydra.exe " & Command)
    '    End
    'End If
    MiRuta = App.Path
    If Dir(App.Path & "\pass.sec") = "" Then CrearPassSec App.Path
    
    '
    Call Crea_Grp
    '
    
    LeeMultiConfig
    CargaIdiomas
    
    If HaySeguridad Then
        FrmLogin.Show
    Else
        FrmPrincipal.Show
    End If
End Sub
Public Sub LeeMultiConfig()
    Dim Arch As Integer
    Dim bUF As String
    Dim Buffer As String
    If Dir(App.Path & "\hydramulti.ini") = "" Then
        CreaConfig_Defecto
        EscribeConfig
        CargaIdiomas
        FrmPrincipal.Show
        MostrarModal FrmPrincipal, FrmConfiguracion
        FrmPrincipal.CambiarIdioma
    Else
        Arch = FreeFile()
        Open App.Path & "\hydramulti.ini" For Input As #Arch
        '***************************
        ' seguridad
        Input #Arch, Buffer, bUF
        Buffer = DesEncripta(Buffer)
        If Not IsNumeric(Left(Buffer, 8)) Then
            MsgBox CargaCadena(851), vbCritical
            End
        Else
            If (Val(Left(Buffer, 4)) + 3) = Val(Mid(Buffer, 5, 4)) Then
                HaySeguridad = False
            Else
                If (Val(Left(Buffer, 4)) + 1) = Val(Mid(Buffer, 5, 4)) Then
                    HaySeguridad = True
                Else
                    MsgBox CargaCadena(851), vbCritical
                    End
                End If
            End If
        End If
        '***************************
        ' idioma
        Input #Arch, ID, bUF
        '***************************
        ' exportación
        Input #Arch, Backup_Path, bUF
        Input #Arch, Buffer, bUF
        Exp_Fecha = Val(Buffer)
        Input #Arch, Buffer, bUF
        Exp_Hora = Val(Buffer)
        Input #Arch, Exp_Positivo, bUF
        Input #Arch, Exp_Negativo, bUF
        Input #Arch, Buffer, bUF
        Exp_Absoluto = Buffer
        Input #Arch, Buffer, bUF
        Exp_Diario = Buffer
        Close #Arch
    End If
End Sub
Public Sub EscribeConfig()
    Dim Arch As Integer
    Dim MiValor As Integer
    Dim MiValor2 As Integer
    Arch = FreeFile()
    Open App.Path & "\hydramulti.ini" For Output As #Arch
    Randomize Time
    MiValor = Int((4500 * Rnd) + 1)
    If HaySeguridad Then
        MiValor2 = MiValor + 1
    Else
        MiValor2 = MiValor + 3
    End If
    Write #Arch, Encripta(Format(MiValor, "0000") & Format(MiValor2, "0000")), "NEVER EDIT NOR REMOVE THIS LINE"
    Write #Arch, ID, "Language / Idioma"
    Write #Arch, Backup_Path, "Backup Path / Ruta para copias de seguridad"
    Write #Arch, Exp_Fecha, "Export : Date Format / Formato de Fecha"
    Write #Arch, Exp_Hora, "Export : Time Format / Formato de Hora"
    Write #Arch, Exp_Positivo, "Export : positive symbol / Signo Positivo"
    Write #Arch, Exp_Negativo, "Export : negative symbol / Signo Negativo"
    Write #Arch, Exp_Absoluto, "Export : Use Absolute Values / Usar valores Absolutos"
    Write #Arch, Exp_Diario, "Export : Only Data from current Date / Sólo datos del día en curso"
    Close #Arch
End Sub
Public Sub CreaConfig_Defecto()
    ID = 0
    HaySeguridad = False
    Backup_Path = App.Path & "\backup"
    Exp_Diario = False
    Exp_Absoluto = False
    Exp_Positivo = "0"
    Exp_Negativo = "1"
    Exp_Fecha = 0
    Exp_Hora = 0
End Sub


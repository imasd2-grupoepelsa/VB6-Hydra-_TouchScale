Attribute VB_Name = "Capa0"
Option Explicit
'///////////////////////////////////////////
'// Este módulo rige las comunicaciones   //
'// a más bajo nivel :                    //
'// controla el envío y recepción de      //
'// paquetes por RS-232 y TCP/IP          //
'///////////////////////////////////////////
Public ScaleType As Long
'*****************************
' Estado
'*****************************
Public Estado As Long
' Estado :
' 0--> No configurado
' 1--> Configurado TCP/IP
'****************************
' errores
' 0 --> OK
' 1 --> Sistema No Configurado
' 2 --> Ocupado con otra tarea
Public Const MAX_MULTI_SOCKETS = 18
' Buffer de envío
Public buffersend As String
'****************************
' Buffer de recepción
Public BufferRec As String
' cuando está trabajando
Public Ocupado As Boolean
'*****************************
Public MyConnObj As New ADODB.Connection
Public MyRecSet As New ADODB.Recordset
Public strStream As New ADODB.Stream
Public sAddIp As String
''''''''''''''''''''''''''''''''''''''''

Public Sub Capa0_Reset_Touch()
    On Error Resume Next
    Estado = 0
    Ocupado = False
    If MyConnObj.State Then
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
    On Error GoTo 0
    CadenadeLog "Reset del OCX completado"
End Sub

Public Sub Capa0_Reset()
    If ScaleType = 11 Then
        Call Capa0_Reset_Touch
    End If
    CadenadeLog "Reset del OCX completado"
End Sub

'***********************
' MachineID : 11
'***********************
'***********************
' TypeConf : Siempre 0
'***********************
' devuelve :
' 0 --> OK, configurado
' 3 --> TCP/IP incorrecto
Public Function Capa0_Configure_Touch(ByVal MachineID As Long, ByVal TypeConf As Long, sConfigure As String) As Long
Dim bytes(5) As String
Dim Contador As Integer
Dim Bucle As Integer
Dim Resp As Boolean
Dim Puerto As Long

    ScaleType = MachineID
    
    If MyConnObj.State Then
        CadenadeLog "Estado de la conexión:" & CStr(MyConnObj.State) & " Se cierra conexión Existente previamente."
        On Error Resume Next
        MyConnObj.Close
        Set MyConnObj = Nothing
        If Err.Number <> 0 Then
            CadenadeLog "Error cerrando conexión previamente abierta:" & CStr(Err.Number) & " descr.:" & Err.Description
        End If
        On Error GoTo 0
    End If
    
    sAddIp = ""
    For Contador = 0 To 4
        bytes(Contador) = ""
    Next Contador
    Contador = 0
    Resp = False
    For Bucle = 1 To Len(sConfigure)
        Select Case Asc(Mid(sConfigure, Bucle, 1))
            Case 48 To 57
                bytes(Contador) = bytes(Contador) & Mid(sConfigure, Bucle, 1)
            Case Asc(".")
                Contador = Contador + 1
            Case Asc(":")
                Contador = 4
            Case Else
                Exit For
        End Select
        If Contador > 4 Then
            Exit For
        End If
    Next Bucle
    If bytes(0) <> "" And bytes(1) <> "" And bytes(2) <> "" And bytes(3) <> "" Then
        bytes(0) = Val(bytes(0))
        bytes(1) = Val(bytes(1))
        bytes(2) = Val(bytes(2))
        bytes(3) = Val(bytes(3))
        If Val(bytes(0)) < 256 And Val(bytes(1)) < 256 And Val(bytes(2)) < 256 And Val(bytes(3)) < 256 Then
            If bytes(4) <> "" Then
                Puerto = bytes(4)
            Else
                Puerto = 6000
            End If
            If Puerto >= 1 And Puerto <= 65535 Then
                sAddIp = CStr(bytes(0)) & "." & CStr(bytes(1)) & "." & CStr(bytes(2)) & "." & CStr(bytes(3))
                Resp = True
            End If
        Else
            Resp = False
        End If
    Else
        Resp = False
    End If
    If sAddIp <> "" And Resp = True Then
        On Error Resume Next
        Resp = False
        CadenadeLog "Cadena de conexión(Uid y Pass no reales): DRIVER={MySQL ODBC 5.1 Driver};Uid=X;password=Y;DATABASE=pcscale;Data Source=" & sAddIp
        MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        If Err.Number <> 0 Then
            
            CadenadeLog "Error: " & Err.Number & " Descr.: " & Err.Description
            
            On Error GoTo 0
            Resp = False
            On Error Resume Next
            MyConnObj.Close
            Set MyConnObj = Nothing
            On Error GoTo 0
            Estado = 0
        Else
            Resp = True
            Estado = 1
        End If
        On Error GoTo 0
    End If
    If Resp Then
        Capa0_Configure_Touch = 0
        CadenadeLog "Conexión realizada con:" & sAddIp
    Else
        Capa0_Configure_Touch = 3
    End If
End Function

Public Function Capa0_Configure(ByVal MachineID As Long, ByVal TypeConf As Long, sConfigure As String) As Long
    Dim MyAux As String
    
    If MachineID = 11 Then
        MyAux = sConfigure
        Capa0_Configure = Capa0_Configure_Touch(11, 0, MyAux)
    End If
End Function


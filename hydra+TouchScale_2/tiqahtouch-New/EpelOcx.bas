Attribute VB_Name = "EpelOcx"
Option Explicit

Public MyConnObj As New ADODB.Connection
Public MyRecSet As New ADODB.Recordset
Public Estado As Long

Public Function Get_Status() As Long
    Get_Status = MyConnObj.State
End Function


Public Function Configure(ByVal MachineID As Long, ByVal TypeConf As Long, ByVal sConfigure As String) As Long
    Dim Resp As Long
    If Dir(App.Path & "\marcatime") <> "" Then
        lMarcaTime = True
    Else
        lMarcaTime = False
    End If
    CadenadeLog "Recibida orden de configuración.Máquina : " & MachineID & ".Tipo conexión : " & TypeConf & ".Cadena Conexión : " & sConfigure
    Resp = Capa0_Configure_Touch(MachineID, TypeConf, sConfigure)
    Configure = Resp
End Function

Public Sub reset()
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
End Sub

Public Function Capa0_Configure_Touch(ByVal MachineID As Long, ByVal TypeConf As Long, sConfigure As String) As Long
Dim Bytes(5) As String
Dim Contador As Integer
Dim Bucle As Integer
Dim Resp As Boolean
Dim Puerto As Long
Dim sIp As String
Dim nResp As Integer
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
    
    sIp = ""
    For Contador = 0 To 4
        Bytes(Contador) = ""
    Next Contador
    Contador = 0
    Resp = False
    For Bucle = 1 To Len(sConfigure)
        Select Case Asc(Mid(sConfigure, Bucle, 1))
            Case 48 To 57
                Bytes(Contador) = Bytes(Contador) & Mid(sConfigure, Bucle, 1)
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
    If Bytes(0) <> "" And Bytes(1) <> "" And Bytes(2) <> "" And Bytes(3) <> "" Then
        Bytes(0) = Val(Bytes(0))
        Bytes(1) = Val(Bytes(1))
        Bytes(2) = Val(Bytes(2))
        Bytes(3) = Val(Bytes(3))
        If Val(Bytes(0)) < 256 And Val(Bytes(1)) < 256 And Val(Bytes(2)) < 256 And Val(Bytes(3)) < 256 Then
            If Bytes(4) <> "" Then
                Puerto = Bytes(4)
            Else
                Puerto = 3306
            End If
            sPortBal = CStr(Puerto)
            If Puerto >= 1 And Puerto <= 65535 Then
                sIp = CStr(Bytes(0)) & "." & CStr(Bytes(1)) & "." & CStr(Bytes(2)) & "." & CStr(Bytes(3))
                Resp = True
            End If
        Else
            Resp = False
        End If
    Else
        Resp = False
    End If
    If sIp <> "" And Resp = True Then
        On Error Resume Next
        
        nResp = ElPing(sIp)
        If nResp = 1 Then
        
            Resp = False
            CadenadeLog "Cadena de conexión(Uid y Pass no reales): DRIVER={MySQL ODBC 5.1 Driver};Uid=X;password=Y;DATABASE=pcscale;Data Source=" & sIp
            If Dir(App.Path & "\odbc53.txt") <> "" Then
                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sIp & "; Port=" & CStr(Puerto) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
            Else
                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sIp & "; Port=" & CStr(Puerto) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
            End If
            
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
        
        Else
            CadenadeLog "Sin ping con:" & sIp
            Resp = False
        End If
    End If
    If Resp Then
        Capa0_Configure_Touch = 0
        CadenadeLog "Conexión realizada con:" & sIp
    Else
        Capa0_Configure_Touch = 3
    End If
End Function

Public Function getNtotalTS(ByVal sfecha As Variant) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    On Error Resume Next
    'frmControl.Switch_mode 0
    MyRecAux.CursorLocation = adUseClient
    MyRecAux.Open "Select * from pcscale.totaldaily where date='" & Format(CDate(sfecha), "yyyy-mm-dd") & "'", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        cancelar = 1
        'Call frmControl.Balanza_ComError(0)
        Resp = 0
    Else
        If MyRecAux.EOF Then
            Resp = 0
        Else
            Resp = MyRecAux.Fields("number")
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    'MyConnObj.Close
    'Set MyConnObj = Nothing
    getNtotalTS = Resp
End Function

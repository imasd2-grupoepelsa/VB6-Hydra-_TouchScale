Attribute VB_Name = "comu2"
Public Sub enviafast(que_enviar, que_llega As String)
Dim config_ini_2
Dim linea_config_2
Dim configura_2(1)
Dim vel_2
Dim puer_2
Dim recibido As Integer
Dim vueltas As Integer
      

'''''''''''''''''''''''''''''Configura la velocidad y los puertos''''
   

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
'        frmConfiguracion.MSComm1.PortOpen = True
        frmConfiguracion.MSComm1.InputLen = 0
        que_llega = Empty
        'recibido = Empty
        
        frmConfiguracion.MSComm1.CommPort = puerto
       
        frmConfiguracion.MSComm1.Output = que_enviar
    
 
        frmConfiguracion.Tmcomu.Interval = 5000
   
        Form1.control.Enabled = True
        
        terMinoes = 0
        
            Do
                DoEvents
                
                If terMinoes = 1 And frmConfiguracion.MSComm1.InBufferCount = 0 And vueltas <= 3 And que_llega = "" Then

                vueltas = vueltas + 1
                
                terMinoes = 0
                                                
'
                End If
'
                If vueltas = 3 Then
'
                vueltas = 0
                frmespere.Caption = LoadResString(id + 354)
                Exit Do
                End If
        
                Loop Until frmConfiguracion.MSComm1.InBufferCount > 0
                               
                                               
                que_llega = que_llega & frmConfiguracion.MSComm1.Input
                
      Form1.control.Enabled = False
   
       If que_llega = "" Then

       'frmARTICULO.Caption = "No hay respuesta"
       frmespere.Caption = LoadResString(id + 354)
       frmmonitor.Caption = LoadResString(id + 354)
       Else
       vueltas = 0
       
       End If
       
        
End Sub

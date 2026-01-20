Attribute VB_Name = "Modulo_FTP"
Option Explicit
' Declaraciones Apis
'***************************************************************
Private Declare Function InternetCloseHandle _
  Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Establece una conexión a internet para poder iniciar seción Ftp
Private Declare Function InternetConnect _
  Lib "wininet.dll" Alias "InternetConnectA" _
  (ByVal hInternetSession As Long, _
   ByVal sServerName As String, _
   ByVal nServerPort As Integer, _
   ByVal sUserName As String, _
   ByVal sPassword As String, _
   ByVal lService As Long, ByVal lFlags As Long, _
   ByVal lContext As Long) As Long
'Conecta al Ftp
Private Declare Function InternetOpen _
  Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, _
    ByVal sProxyName As String, _
    ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long
'Recupera un archivo
Private Declare Function FtpGetFile Lib "wininet.dll" _
   Alias "FtpGetFileA" (ByVal hConnect As Long, _
  ByVal lpszRemoteFile As String, _
  ByVal lpszNewFile As String, ByVal fFailIfExists As Long, _
  ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
  ByRef dwContext As Long) As Boolean
'Escribe un archivo
Private Declare Function FtpPutFile Lib "wininet.dll" _
   Alias "FtpPutFileA" (ByVal hConnect As Long, _
  ByVal lpszLocalFile As String, _
  ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
  ByVal dwContext As Long) As Boolean
  
Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String) As Boolean

'Api Para los errores
Private Declare Function InternetGetLastResponseInfo _
  Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
  (lpdwError As Long, ByVal lpszBuffer As String, _
   lpdwBufferLength As Long) As Boolean
'***************************************************
'str
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer
'***************************************
Public Declare Function FtpGetFileSize Lib "wininet.dll" Alias "FtpGetFileSizeA" _
(ByVal hFtpSession As Long, ByVal lpszFileSize As Long) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
(ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long

'Type para atributos de fecha y hora de archivos
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const MAX_PATH = 260
  
'Otros atributos de archivo tamaño, nombre, fecha etc..
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'Busca el primer archivo de un path
Private Declare Function FtpFindFirstFile Lib "wininet.dll" _
   Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, _
   ByVal lpszSearchFile As String, _
   lpFindFileData As WIN32_FIND_DATA, _
   ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const MAXDWORD As Double = (2 ^ 32) - 1

Public FTP_HandleConect(2) As Long
Public FTP_hOpen(2) As Long
Public nRetConnFTP As Long

'retorno 0 no ha sido posible...
Public Function Conectar_FTP(ByVal sIpServidor As String, ByVal sUser As String, ByVal sPass As String, ByVal nIndex As Integer) As Long
    Dim Ret As Long
    Dim nResp As Integer
    
    Ret = 0
    FTP_hOpen(nIndex) = 0
    FTP_HandleConect(nIndex) = 0
    nResp = ElPing(sIpServidor)
    If nResp = 1 Then
    
        FTP_hOpen(nIndex) = InternetOpen(vbNullString, 0, vbNullString, vbNullString, 0)
        If FTP_hOpen(nIndex) <> 0 Then
            FTP_HandleConect(nIndex) = InternetConnect(FTP_hOpen(nIndex), sIpServidor, 21, sUser, sPass, 1, IIf(True, &H8000000, 0), 0)
            Ret = InternetSetOptionStr(FTP_HandleConect(nIndex), 2, Format(2000, "0000"), 4)  'timeout conexión
            Ret = InternetSetOptionStr(FTP_HandleConect(nIndex), 5, Format(4000, "0000"), 4)  'timeout envio
            Ret = InternetSetOptionStr(FTP_HandleConect(nIndex), 6, Format(4000, "0000"), 4)  'timeout recepción
        End If
    
    End If
    Conectar_FTP = Ret
End Function
'retorno 0 no ha sido posible...
Public Function Desconectar_FTP(ByVal nIndex As Integer) As Long
    Dim Ret As Long
    Ret = InternetCloseHandle(FTP_HandleConect(nIndex))
    If Ret <> 0 Then
        Ret = InternetCloseHandle(FTP_hOpen(nIndex))
    End If
    Desconectar_FTP = Ret
End Function
'retorno 0 no ha sido posible...
Public Function Enviar_Fichero_FTP(ByVal sOrgFile As String, ByVal sDesFile As String, ByVal nIndex As Integer) As Long
    Dim Ret As Boolean
    Ret = FtpPutFile(FTP_HandleConect(nIndex), sOrgFile, sDesFile, 1, 0)
    If Ret Then
        Enviar_Fichero_FTP = 1
    Else
        Enviar_Fichero_FTP = 0
    End If
    Do_Events
End Function
'retorno 0 no ha sido posible...
Public Function Recibir_Fichero_FTP(ByVal sOrgFile As String, ByVal sDesFile As String, ByVal nIndex As Integer) As Long
    Dim Ret As Boolean
    Ret = FtpGetFile(FTP_HandleConect(nIndex), sOrgFile, sDesFile, False, 0, 1, 0)
    'Ret = FtpGetFile(FTP_HandleConect(nIndex), sOrgFile, sDesFile, True, 0, 1, 0) 'falla si sobreescribirsobreescribir
    If Ret Then
        Recibir_Fichero_FTP = 1
    Else
        Recibir_Fichero_FTP = 0
    End If
    Do_Events
End Function
'retorno 0 no ha sido posible...
Public Function Borrar_Fichero_FTP(ByVal sOrgFile As String, ByVal nIndex As Integer) As Long
    Dim Ret As Boolean
    Ret = FtpDeleteFile(FTP_HandleConect(nIndex), sOrgFile)
    If Ret Then
        Borrar_Fichero_FTP = 1
    Else
        Borrar_Fichero_FTP = 0
    End If
    Do_Events
End Function

'''''''''''''''''
Public Function conecta_FTP() As Long
Dim Ret As Long
Dim nFich As Integer

    If nRetConnFTP <> 0 Then
        Call Cierra_FTP
    End If
        
    If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
        Dim siP, susR, spsS
        nFich = FreeFile()
        Open App.Path & "\getputfileftp.dat" For Input As #nFich
        'Do While Not EOF(nFich)
            Line Input #nFich, siP
            Line Input #nFich, susR
            Line Input #nFich, spsS
        'Loop
        Close #nFich
        Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
        Ret = Err.LastDllError
        If Ret <> 0 Then 'error en DLL
            Ret = 0
        Else
            Ret = 1
        End If
    Else
        Ret = 0
    End If
    nRetConnFTP = Ret
    conecta_FTP = Ret
End Function
Public Sub Cierra_FTP()
Dim Ret As Long
    Do_Events
    Ret = Desconectar_FTP(0)
    Do_Events
    FTP_hOpen(0) = 0
    FTP_HandleConect(0) = 0
    nRetConnFTP = 0
End Sub

Public Function getfilesize(ByVal sFile As String, ByVal nIndex As Integer) As Long
Dim hFile As Long
Dim filesize As Long
Dim nFileSizeLow As Long
Dim nFileSizeHigh As Long
Dim sfilesize As String
Dim pData As WIN32_FIND_DATA
Dim Ret As Long
    
    'crea un buffer
    pData.cFileName = String(MAX_PATH, 0)
    'busca el primer fichero
    hFile = FtpFindFirstFile(FTP_HandleConect(nIndex), sFile, pData, 0, 0)
    
    'hFile = FtpOpenFile(FTP_HandleConect(nIndex), sFile, &H80000000, 1, 0)
    'hFile = hFile
    'hFile = FtpOpenFile(FTP_hOpen(nIndex), sFile, &H80000000, 1, 0)
    If hFile = 0 Then 'No existe
        filesize = 0
        GoTo SalFileSize
    End If
    'nFileSizeLow = FtpGetFileSize(hFile, nFileSizeHigh)
    filesize = (pData.nFileSizeHigh * (MAXDWORD + 1) + pData.nFileSizeLow)
SalFileSize:
    Ret = InternetCloseHandle(hFile)
    getfilesize = filesize
End Function



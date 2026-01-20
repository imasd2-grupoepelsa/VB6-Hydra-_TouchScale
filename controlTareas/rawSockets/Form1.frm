VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code will send a raw tcpip packet without data using a dll
' just put rawtcpip.dll in windows\system32 and run this code.
' you need to pass 11 arguments. The dll is raw with out much error handling.
' I will try to include some error handling in it.
' Arguments to pass are all type string.
' Argument list is : Destination IP, Destination Port.
' Source IP,Source Port, TTL, Flags(URG,ACK,PSH,RST,SYN,FIN)
' it will return 0 in successfull packet transmission.
' please send me bugs reports and comments on prashsax@yahoo.com

Private Declare Function sendpacket Lib "rawtcpip" (ByVal dip As String, ByVal dp As String, ByVal sip As String, ByVal sp As String, ByVal ttl As String, ByVal urg As String, ByVal ack As String, ByVal psh As String, ByVal rst As String, ByVal syn As String, ByVal fin As String) As Long

Private Sub Command1_Click()
retval = sendpacket("58.56.32.56", "139", "10.19.44.55", "1545", "128", "", "", "", "", "syn", "")
MsgBox Str(retval)
End Sub


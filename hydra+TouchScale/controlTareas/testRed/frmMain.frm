VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Test OSICMP"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTraceroute 
      Caption         =   "Traceroute"
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtHost 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPing_Click()
  Dim oPing As New OSICMP.Ping
  Dim i As Integer 'request counter
  Dim s As String
  
On Error GoTo err_handler:
 
  For i = 0 To 3
    oPing.Send txtHost.Text
    
    If i = 0 Then
      If txtHost.Text <> oPing.IP Then s = txtHost.Text & " [" & oPing.IP & "]" Else s = txtHost.Text
      txtOutput.Text = "Pinging " & s & " with " & oPing.PacketSize & " bytes of data:" & vbCrLf & vbCrLf
    End If
    
    txtOutput.Text = txtOutput.Text & "Reply from " & _
      oPing.IP & ": bytes=" & oPing.PacketSize & " time=" & _
      oPing.RoundTripTime & "ms TTL=" & oPing.TTL & vbCrLf
    txtOutput.Refresh
    
    oPing.Sleep 1000
  Next
  
  txtOutput.Text = txtOutput.Text & vbCrLf & "complete"
  Exit Sub
  
err_handler:
  txtOutput.Text = "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub cmdTraceroute_Click()
  Dim oTraceroute As New OSICMP.Traceroute
  Dim i As Integer 'request counter
  Dim j As Integer 'hop counter
  Dim s As String
  Dim bTimeout As Boolean
  Dim sHopIP As String
  
On Error GoTo err_handler:
  txtOutput.Text = ""
  For j = 1 To 30 'limit tracing to 30 hops
    sHopIP = "request timed out"
    For i = 1 To 3 '3 requests per hop
      bTimeout = False
      oTraceroute.Send txtHost.Text, , , j
      
      If i = 1 Then
        If j = 1 Then
          If txtHost.Text <> oTraceroute.IP Then s = txtHost.Text & " [" & oTraceroute.IP & "]" Else s = txtHost.Text
          txtOutput.Text = "Tracing route to " & s & vbCrLf & _
            "over a maximum of 30 hops:" & vbCrLf & vbCrLf
        End If
        txtOutput.Text = txtOutput.Text & j & "    "
      End If
      
      If bTimeout Then
        txtOutput.Text = txtOutput.Text & "   *   "
      Else
        txtOutput.Text = txtOutput.Text & "    " & oTraceroute.RoundTripTime & " ms"
        If sHopIP = "request timed out" Then sHopIP = oTraceroute.HopIP
      End If
      txtOutput.Refresh
    Next
    txtOutput.Text = txtOutput.Text & "  " & sHopIP & vbCrLf
    If oTraceroute.IP = sHopIP Then
      txtOutput.Text = txtOutput.Text & vbCrLf & "Trace complete." & vbCrLf
      Exit For
    End If
    oTraceroute.Sleep 1000
  Next
  
  If oTraceroute.IP <> sHopIP Then _
    txtOutput.Text = txtOutput.Text & vbCrLf & "Host unreachable." & vbCrLf
  Exit Sub
  
err_handler:
  If Err.Number = 11010 Then 'timeout
    bTimeout = True
    Err.Clear
    Resume Next
  End If
  
  txtOutput.Text = "Error " & Err.Number & ": " & Err.Description
End Sub

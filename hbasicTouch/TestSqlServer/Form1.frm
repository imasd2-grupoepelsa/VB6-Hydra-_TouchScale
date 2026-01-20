VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1560
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    
    'MyConnObj.Open "Provider=sqloledb;Data Source=10.8.1.32;Uid=sa;pwd=;Initial Catalog=HiOScale;integrated security=SSPI;persist security info=True;"
    'MyConnObj.Open "Provider=sqloledb.1;Data Source=10.8.1.32;Initial Catalog=HiOScale;Uid=sa;password=;"
    On Error Resume Next
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=123;Initial Catalog=HiOScale;Data Source=10.8.1.108\SQLEXPRESS"
    If Err.Number = 0 Then
        MsgBox "Connect with:" & "Provider=SQLOLEDB;Uid=sa;password=123;Initial Catalog=HiOScale;Data Source=10.8.1.108\SQLEXPRESS"
        MyConnObj.Close
    Else
        On Error GoTo 0
        MsgBox "Connect error with:" & "Provider=SQLOLEDB;Uid=sa;password=123;Initial Catalog=HiOScale;Data Source=10.8.1.32;"
    End If
    'On Error Resume Next
    'MyConnObj.Open "Provider=sqloledb;Data Source=10.8.1.108;Initial Catalog=HiOScale;integrated security=SSPI;persist security info=True;"
    'If Err.Number = 0 Then
    '    MsgBox "Connect with:" & "Provider=sqloledb;Data Source=10.8.1.32;Initial Catalog=HiOScale;integrated security=SSPI;persist security info=True;"
    '    MyConnObj.Close
    'Else
    '   On Error GoTo 0
    '    MsgBox "Connect error with:" & "Provider=sqloledb;Data Source=10.8.1.32;Initial Catalog=HiOScale;integrated security=SSPI;persist security info=True;"
    'End If

End Sub

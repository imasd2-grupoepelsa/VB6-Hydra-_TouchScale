Attribute VB_Name = "MultiStore_Link"
Option Explicit
Public Type TTienda_Data
    numero As Long
    Path As String

End Type
Public Sub Dame_Tiendas(Tiendas() As TTienda_Data, NTiendas As Long)
    Dim sPath As String
    Dim MyPath As String
    Dim bucle As Long
    Dim Posibles() As Long
    Dim NPosibles As Long
    NTiendas = 0
    sPath = Dir(App.Path & "\T*", vbDirectory)
    If sPath <> "" Then
        Do While sPath <> ""

            If UCase(Mid(sPath, 1, 1)) = "T" Then
                If IsNumeric(Mid(sPath, 2, 3)) Then
                    If Val(Mid(sPath, 2, 3)) <= 255 Then
                        NPosibles = NPosibles + 1
                        ReDim Preserve Posibles(NPosibles)
                        Posibles(NPosibles - 1) = Val(Mid(sPath, 2, 3))
                    End If
                End If
            End If

            sPath = Dir()
        Loop
    End If
    If NPosibles > 0 Then
        For bucle = 0 To NPosibles
            MyPath = App.Path & "\t" & Format(Posibles(bucle), "00")
            If Dir(MyPath, vbDirectory) <> "" Then
                If Dir(MyPath & "\dbasetouch.mdb") <> "" Then
                    If Dir(MyPath & "\hydratouch.ini") <> "" Then
                        NTiendas = NTiendas + 1
                        ReDim Preserve Tiendas(NTiendas)
                        Tiendas(NTiendas - 1).Path = MyPath
                        Tiendas(NTiendas - 1).numero = Posibles(bucle)
                    End If
                End If
            End If
        Next bucle
    End If
End Sub
Public Function ArticuloCompatible(ByVal nTienda As Long, ByVal nCod As Long, ByVal nPlu As Long, ByVal nMos As Long) As Long
' articulo compatible:
' 0 --> No existe
' 1 --> OK
' 2 --> Plu incorrecto
' 3 --> Mostrador incorrecto
' 4 --> Codigo incorrecto
    Dim sPath As String
    Dim Resp As Long
    Dim Registro As RecordNet
    Dim BufPath As String
    Resp = 0
    BufPath = Base_General
    Base_General = App.Path & "\t" & Format(nTienda, "00") & "\dbasetouch.mdb"
    Set Registro = New RecordNet
    Registro.OpenRecordset "select codigo,plu,codi_ident from articulo where borrado=false and codigo=" & nCod
    If Registro.EOF Then
        Registro.OpenRecordset "select codigo,plu,codi_ident from articulo where borrado=false and codi_ident=" & nMos & " and plu=" & nPlu
        If Not Registro.EOF Then Resp = 4
    Else
        Resp = 1
        If nPlu <> Registro.Fields("plu") Then Resp = 2
        If nMos <> Registro.Fields("codi_ident") Then Resp = 3
    End If
    Base_General = BufPath
    ArticuloCompatible = Resp
End Function


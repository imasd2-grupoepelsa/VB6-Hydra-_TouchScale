Attribute VB_Name = "Idiomas"
Option Explicit
Public nIdioma As Long
Public sIdioma() As String
Public Sub CargaIdiomas()
    Dim Base As Database
    Dim Registro As Recordset
    Dim Bucle As Long
    Dim MyID As Integer
    MyID = ID + 1
    Set Base = OpenDatabase(App.Path & "\langtouch\langtouch.mdb")
    Set Registro = Base.OpenRecordset("select max(id) from texts")
    With Registro
        nIdioma = Val(.Fields(0))
        ReDim Preserve sIdioma(nIdioma + 1)
    End With
    Set Registro = Base.OpenRecordset("select * from texts order by id")
    With Registro
        .MoveFirst
        Do Until .EOF
            If Not IsNull(.Fields("text" & Format(MyID, "0"))) Then
                sIdioma(.Fields("id")) = .Fields("text" & Format(MyID, "0"))
            End If
            .MoveNext
        Loop
    End With
    Base.Close
End Sub
Public Function CargaCadena(ByVal NumCad As Integer) As String
    If NumCad < nIdioma Then
        CargaCadena = sIdioma(NumCad)
    Else
        CargaCadena = "Cad : " & NumCad
    End If
End Function

Attribute VB_Name = "Envios_Ahorramas"
Option Explicit
Public Sub Exportar_VAcuno()
' Formato de la exportación :
' 1 caracter modo (local-->1, venta -->0)
' 3 caracteres Número de ficha(1-255)
' 10 caracteres fecha dd/mm/aaaa
' 5 caracteres nº de operaciones
' 5 caracteres nº de unidades
' 10 caracteres Importe
' 10 caracteres peso (en gramos)
    Dim Fichero As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Buffer As String
    Fichero = FreeFile
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from gtvacuno")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Open Miruta & "\totficha.vac" For Append As Fichero
            Do Until .EOF
                Buffer = "0"
                Buffer = Buffer & Format(!ficha + 1, "000")
                Buffer = Buffer & Format(!Fecha, "dd/mm/yyyy")
                Buffer = Buffer & Format(!vnumoperaciones, "00000")
                Buffer = Buffer & Format(!vnumunidades, "00000")
                Buffer = Buffer & Format(!vimporte, "0000000000")
                Buffer = Buffer & Format(!vpeso * 1000, "0000000000")
                Print #Fichero, Buffer
                Buffer = "1"
                Buffer = Buffer & Format(!ficha + 1, "000")
                Buffer = Buffer & Format(!Fecha, "dd/mm/yyyy")
                Buffer = Buffer & Format(!lnumoperaciones, "00000")
                Buffer = Buffer & Format(!lnumunidades, "00000")
                Buffer = Buffer & Format(!limporte, "0000000000")
                Buffer = Buffer & Format(!lpeso * 1000, "0000000000")
                Print #Fichero, Buffer
                .Movenext
            Loop
            Close #Fichero
        End If
    End With
    '1.7.2
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Registro = Nothing
    Set Base = Nothing
    'Workspaces(0).close

End Sub


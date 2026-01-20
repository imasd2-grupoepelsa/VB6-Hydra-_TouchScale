Attribute VB_Name = "XMLTools"
Option Explicit
Public Function DB_2XML(ByVal sFiltro As String, ByVal sPath As String)
    Dim Bucle As Long
    Dim B2 As Long
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Dim Arch As New ClsXMLWriter
    Dim nCan As Long
    Dim sAttrib(1) As String
    Dim sValue(1) As String
    On Error Resume Next
    Arch.document = sPath
    Arch.OpenDoc "hydra"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Ruta no válida", vbCritical
    End If
    On Error GoTo 0
    Set Base = OpenDatabase(Base_General)
    For Bucle = 0 To Base.TableDefs.Count - 1
        If UCase(left(Base.TableDefs(Bucle).Name, 4)) <> "MSYS" Then
            Set Registro = Base.OpenRecordset(Base.TableDefs(Bucle).Name)
            With Registro
                If Not .EOF Then
                    .MoveFirst
                    
                    Do Until .EOF
                        Arch.SetSimpleBlock Base.TableDefs(Bucle).Name
                        For B2 = 0 To .Fields.Count - 1
                            If IsNull(.Fields(B2)) Then
                                Select Case .Fields(B2).Type
                                    Case dbText
                                        sValue(0) = ""
                                    Case dbDate
                                        sValue(0) = "01/01/01 00:00:00"
                                    Case dbBoolean
                                        sValue(0) = "FALSE"
                                    Case Else
                                        sValue(0) = "0"
                                End Select
                            Else
                                Select Case .Fields(B2).Type
                                    Case dbText

                                        sValue(0) = Replace(.Fields(B2), Chr(34), "'")
                                    Case dbDate
                                        sValue(0) = Format(.Fields(B2), "DD/MM/YY hh:mm:ss")
                                    Case dbBoolean
                                        If .Fields(B2) = False Then
                                            sValue(0) = "FALSE"
                                        Else
                                            sValue(0) = "TRUE"
                                        End If
                                    Case Else
                                        sValue(0) = CStr(.Fields(B2))
                                End Select
                            End If
                            sAttrib(0) = "value"
                            Arch.SetDataBlock .Fields(B2).Name, 1, sAttrib, sValue
                            Arch.EndDataBlock
                        Next B2
                        Arch.EndDataBlock
                        .Movenext
                    Loop
                    Arch.EndDataBlock
                End If
            End With
        End If
    Next Bucle
    Base.Close
End Function

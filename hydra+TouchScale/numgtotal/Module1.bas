Attribute VB_Name = "Module1"
Option Explicit
Public lTiendas(99) As Boolean
Public sPathBusca(99) As String
Public nTiendas As Integer
Public Type dBtiendas
    nDbs(99) As Integer
    sNomDbs(99) As String
End Type
Dim tiendas As dBtiendas



Public Sub main()
    Dim nCont, nCont1  As Integer
    Dim nInicio As Integer
    Dim nFin    As Integer
    
    If Dir(App.Path & "\hydramulti.exe") = "" Then
        nTiendas = 0
        nInicio = 0
        nFin = 0
        sPathBusca(0) = App.Path & "\s"
    Else
        nTiendas = 99
        nInicio = 1
        nFin = 99
        For nCont = 1 To 99
            If Dir(App.Path & "\t" & Format(nCont, "00"), vbDirectory) <> "" Then
                lTiendas(nCont) = True
                sPathBusca(nCont) = App.Path & "\t" & Format(nCont, "00") & "\s"
            Else
                lTiendas(nCont) = False
                sPathBusca(nCont) = ""
            End If
        Next nCont
    End If
    

    For nCont = nInicio To nFin
        frmGT.FileL.Path = sPathBusca(nCont)
        frmGT.FileL.Pattern = "*.mdb"
        tiendas.nDbs(nCont) = frmGT.FileL.ListCount
        If frmGT.FileL.ListCount > 0 Then
            For nCont1 = 0 To frmGT.FileL.ListCount - 1
                frmGT.FileL.ListIndex = nCont1
                tiendas.sNomDbs(nCont1) = frmGT.FileL.FileName
            
            Next nCont1
        
        End If
    Next nCont
    
    Load frmGT
    frmGT.Visible = True
    
    
End Sub

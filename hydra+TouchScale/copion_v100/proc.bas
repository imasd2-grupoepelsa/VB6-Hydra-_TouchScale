Attribute VB_Name = "proc"
Option Explicit
'''''''''''''''
'''''''''''''''
Public Function backup(Optional ByVal lSi As Boolean) As Integer
    Dim fecha_ini As Date
    Dim fecha_fin As Date
    Dim cnt As Long
    Dim MaxT As Long
    Dim dbSYS As DAO.Database
    Dim rstSYS As DAO.Recordset
    Dim sTabla As String
    Dim sFecha As String
    Dim sSQL As String
    Dim smm As String
    Dim syyyymm As String
    Dim pntRini As Long
    Dim pntRfin As Long
    Dim res As Integer
    Dim cntr As Long
    Dim n As Integer
    
    On Error GoTo ret
    res = 0
    
    'fecha_ini = Format("01/01/00", "mm/dd/yy")
    'fecha_fin = Format((Now - DiasBorrado), "mm/dd/yy")
    
    
    fecha_ini = Format("01/01/00", "mm/dd/yy")
    If lSi Then
        fecha_fin = Format((Now), "mm/dd/yy")
    Else
        fecha_fin = Format((Now - DiasBorrado), "mm/dd/yy")
    End If
    
    If UsaFichasVacuno Then
        'MaxT = 11
        MaxT = 9
    Else
        'MaxT = 10
        MaxT = 8
    End If
    
    frmProc.Visible = True
    frmProc.procbar.Min = 0
    frmProc.procbar.Max = MaxT
    Set dbSYS = OpenDatabase(var.DBASE_MDB)
    For cnt = 0 To MaxT
nLOOp:
        
        Select Case cnt
            Case 0
                sTabla = "cabecera"
                sFecha = "D_FECHA"
            Case 1
                cnt = cnt + 1
                GoTo nLOOp
            Case 2
                cnt = cnt + 1
                GoTo nLOOp
            
            Case 3
                sTabla = "gtarti"
                sFecha = "FECHA"
            Case 4
                sTabla = "gthora"
                sFecha = "FECHA"
            Case 5
                sTabla = "gtInfVen"
                sFecha = "FECHA"
            Case 6
                sTabla = "gtsecs"
                sFecha = "FECHA"
            Case 7
                sTabla = "gtvend"
                sFecha = "FECHA"
            Case 8
                sTabla = "tickets"
                sFecha = "D_FECHA"
            'Case 9
            '    cnt = cnt + 1
            '    GoTo nLOOp
            'Case 10
            '    'cnt = cnt + 1
            '    'GoTo nLOOp
            '    sTabla = "cliente"
            '    sFecha = ""
            
            Case 9 '11
                sTabla = "GTVacuno"
                sFecha = "FECHA"
        End Select
        If cnt <> 11 Then
            If cnt = 10 Then
                sSQL = "SELECT *" & _
                       " FROM " & sTabla
            
            Else
                sSQL = "SELECT *" & _
                       " FROM " & sTabla & _
                       " WHERE " & sFecha & " >= #" & fecha_ini & _
                       "# AND " & sFecha & " <= #" & fecha_fin & _
                       "# ORDER BY " & sFecha
            End If
        Else
            sSQL = "SELECT *" & _
                   " FROM " & sTabla & _
                   " WHERE " & sFecha & ">= #" & fecha_ini & _
                   "# AND " & sFecha & "<= #" & fecha_fin & _
                   "# ORDER BY " & sFecha
       
        End If
        Set rstSYS = dbSYS.OpenRecordset(sSQL)
        If (Not rstSYS.EOF) Then
            rstSYS.MoveFirst
            pntRini = 0
            pntRfin = -1
            If cnt <> 10 Then
                smm = Format(rstSYS.Fields(sFecha), "mm")
                syyyymm = Format(rstSYS.Fields(sFecha), "yyyymm")
                Do While (Not rstSYS.EOF)
                    If (smm <> Format(rstSYS.Fields(sFecha), "mm")) Then
                        odb.copia_dbase syyyymm, sTabla, rstSYS, pntRini, pntRfin
                        'pntR = rstSYS.AbsolutePosition + 1
                        pntRini = pntRfin + 1
                        smm = Format(rstSYS.Fields(sFecha), "mm")
                        syyyymm = Format(rstSYS.Fields(sFecha), "yyyymm")
                    Else
                        pntRfin = pntRfin + 1
                        rstSYS.MoveNext
                    End If
                Loop
            
                odb.copia_dbase syyyymm, sTabla, rstSYS, pntRini, pntRfin
                
            Else
                
                'Do While (Not rstSYS.EOF)
                '    'If (smm <> Format(rstSYS.Fields(sFecha), "mm")) Then
                '        odb.copia_dbase syyyymm, sTabla, rstSYS, pntRini, pntRfin
                '    '    'pntR = rstSYS.AbsolutePosition + 1
                '    '    pntRini = pntRfin + 1
                '    '    smm = Format(rstSYS.Fields(sFecha), "mm")
                '    '    syyyymm = Format(rstSYS.Fields(sFecha), "yyyymm")
                '    'Else
                '        pntRfin = pntRfin + 1
                '        rstSYS.MoveNext
                '    'End If
                'Loop
                rstSYS.MoveLast
                pntRini = 1
                pntRfin = rstSYS.RecordCount
                odb.copia_dbase syyyymm, sTabla, rstSYS, pntRini, pntRfin
                
            End If
            
        End If
        rstSYS.Close
        'c2f revisar
        n = FreeFile()
        Open App.Path & "\registro.txt" For Append As #n
        Print #n, sSQL
        
        Set rstSYS = dbSYS.OpenRecordset(sSQL)
        If Not rstSYS.EOF Then
            Print #n, "número de registros:" & Format(rstSYS.RecordCount, "000000")
            With rstSYS
                .MoveFirst
                Do While Not .EOF
                    'Print #n, "Registro:..." & Format(rstSYS.PercentPosition, "000000") & " %"
                    If cnt <> 10 Then
                        .Delete
                    End If
                    .MoveNext
                Loop
            End With
        End If
        Close #n
        ''''''''''''
        Set rstSYS = Nothing
        frmProc.procbar.Value = cnt
        frmProc.procbar.Refresh
    Next cnt
ret:
    frmProc.procbar.Value = frmProc.procbar.Max
    frmProc.procbar.Refresh
    dbSYS.Close
    If (Err.Number <> 0) Then backup = 1
    Err.Clear
    On Error GoTo 0
End Function
''''''''''''
''''''''''''

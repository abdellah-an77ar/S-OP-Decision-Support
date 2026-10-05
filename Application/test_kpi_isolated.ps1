$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

$dbgFile = "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_kpi.txt"
if (Test-Path $dbgFile) { Remove-Item $dbgFile -Force }

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.CodeModule.AddFromString(@"
Private Sub Dbg(msg As String)
    Dim f As Integer: f = FreeFile
    Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_kpi.txt" For Append As #f
    Print #f, Format(Now(), "hh:mm:ss") & " - " & msg
    Close #f
End Sub

Public Sub TraceKPI()
    On Error GoTo EH
    Dbg "Starting TraceKPI"
    Dim wsKpi As Worksheet: Set wsKpi = ThisWorkbook.Sheets(SH_KPI)
    Dim wsGap As Worksheet: Set wsGap = ThisWorkbook.Sheets(SH_GAP)
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets(SH_INVENTORY)
    Dim wsFc  As Worksheet: Set wsFc  = ThisWorkbook.Sheets(SH_FORECAST)
    Dim wsDp  As Worksheet: Set wsDp  = ThisWorkbook.Sheets(SH_DEMAND_PLAN)

    Dbg "Calling ClearSheet SH_KPI"
    ClearSheet SH_KPI
    Dbg "ClearSheet done"

    Dim cycleID As String: cycleID = GetCycleID()
    Dbg "cycleID=" & cycleID
    Dim rowOut As Long: rowOut = 2
    Dim kIdx As Long: kIdx = 1

    Dim totDem As Double: totDem = 0
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dbg "dpLast=" & dpLast
    Dim dr As Long
    For dr = 2 To dpLast
        If IsNumeric(wsDp.Cells(dr, 5).Value) Then totDem = totDem + CDbl(wsDp.Cells(dr, 5).Value)
    Next dr
    Dbg "totDem=" & totDem

    Dim totRup As Double: totRup = 0
    Dim invLast As Long: invLast = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    Dbg "invLast=" & invLast
    Dim ir As Long
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then
            If IsNumeric(wsInv.Cells(ir, 6).Value) Then totRup = totRup + CDbl(wsInv.Cells(ir, 6).Value)
        End If
    Next ir
    Dbg "totRup=" & totRup

    Dim svc As Double: If totDem > 0 Then svc = (totDem - totRup) / totDem * 100 Else svc = 100
    Dbg "svc=" & svc

    Dbg "Calling AddKPI 1"
    AddKPI wsKpi, rowOut, kIdx, "Taux de service global", "Global", "Horizon S.O.P.", svc, "%", 95, 90, cycleID, True
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 1 done"

    Dim maxUtil As Double: maxUtil = 0
    Dim capLast As Long: capLast = wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row
    Dbg "capLast=" & capLast
    Dim cr As Long
    For cr = 2 To capLast
        If IsNumeric(wsCap.Cells(cr, 7).Value) Then
            If CDbl(wsCap.Cells(cr, 7).Value) > maxUtil Then maxUtil = CDbl(wsCap.Cells(cr, 7).Value)
        End If
    Next cr
    Dbg "maxUtil=" & maxUtil
    AddKPI wsKpi, rowOut, kIdx, "Utilisation max capacite", "Capacity", "Peak horizon", maxUtil, "%", 85, 95, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 2 done"

    Dim sumW As Double: Dim nW As Long: nW = 0
    Dim fcLast As Long: fcLast = wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row
    Dbg "fcLast=" & fcLast
    Dim fr As Long
    For fr = 2 To fcLast
        If UCase(CStr(wsFc.Cells(fr, 7).Value)) = "RETENUE" Then
            If IsNumeric(wsFc.Cells(fr, 5).Value) Then
                sumW = sumW + CDbl(wsFc.Cells(fr, 5).Value): nW = nW + 1
            End If
        End If
    Next fr
    Dim wapeMoy As Double: If nW > 0 Then wapeMoy = sumW / nW
    Dbg "wapeMoy=" & wapeMoy
    AddKPI wsKpi, rowOut, kIdx, "WAPE moyen forecast", "Forecast", "Methode retenue", wapeMoy, "%", 20, 30, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 3 done"

    Dim nRup As Long: nRup = 0
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then nRup = nRup + 1
    Next ir
    Dbg "nRup=" & nRup
    AddKPI wsKpi, rowOut, kIdx, "Nombre de ruptures de stock", "Inventory", "Horizon", nRup, "cas", 0, 2, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 4 done"

    Dim nSurstock As Long: nSurstock = 0
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "SURSTOCK" Then nSurstock = nSurstock + 1
    Next ir
    Dbg "nSurstock=" & nSurstock
    AddKPI wsKpi, rowOut, kIdx, "Nombre de surstocks", "Inventory", "Horizon", nSurstock, "cas", 0, 3, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 5 done"

    Dim nSurch As Long: nSurch = 0
    For cr = 2 To capLast
        If CStr(wsCap.Cells(cr, 8).Value) = "SURCHARGE" Then nSurch = nSurch + 1
    Next cr
    Dbg "nSurch=" & nSurch
    AddKPI wsKpi, rowOut, kIdx, "Usines-mois en surcharge", "Capacity", "Horizon", nSurch, "cas", 0, 2, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 6 done"

    Dim totDef As Double: totDef = 0
    Dim gapLast As Long: gapLast = wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row
    Dbg "gapLast=" & gapLast
    For cr = 2 To gapLast
        If IsNumeric(wsGap.Cells(cr, 5).Value) Then
            If CDbl(wsGap.Cells(cr, 5).Value) < 0 Then totDef = totDef + Abs(CDbl(wsGap.Cells(cr, 5).Value))
        End If
    Next cr
    Dbg "totDef=" & totDef
    AddKPI wsKpi, rowOut, kIdx, "Deficit capacite total (unites)", "Gap", "Horizon", totDef, "unites", 0, 1000, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 7 done"

    AddKPI wsKpi, rowOut, kIdx, "Demande totale planifiee", "Demand", "Horizon S.O.P.", totDem, "unites", 0, 0, cycleID, True
    rowOut = rowOut + 1: kIdx = kIdx + 1
    Dbg "AddKPI 8 done"

    Dbg "Adding Table to wsKpi"
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsKpi.ListObjects("tbl_KPI"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsKpi.ListObjects.Add(xlSrcRange, wsKpi.Range("A1:J" & (rowOut - 1)), , xlYes)
            lo.Name = "tbl_KPI": lo.TableStyle = "TableStyleMedium2"
        End If
        wsKpi.Columns("B").ColumnWidth = 32
    End If
    Dbg "Table done. Calling GenererAlertes"

    GenererAlertes wsKpi
    Dbg "GenererAlertes done!"
    Exit Sub
EH:
    Dbg "ERROR: " & Err.Description
End Sub
"@)

    Write-Host "Running TraceKPI..."
    $xl.Run("TraceKPI")
    Write-Host "TraceKPI completed!"

} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

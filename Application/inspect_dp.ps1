$xl = New-Object -ComObject Excel.Application
$wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
$wsDp = $wb.Sheets.Item("DEMAND_PLAN")
$lr = $wsDp.Cells.Item($wsDp.Rows.Count, 1).End(-4162).Row
Write-Host "DEMAND_PLAN Rows: $lr"
for ($r = 1; $r -le [Math]::Min(10, $lr); $r++) {
    Write-Host "Row $r : $($wsDp.Cells.Item($r,1).Text) | $($wsDp.Cells.Item($r,2).Text) | $($wsDp.Cells.Item($r,3).Text) | $($wsDp.Cells.Item($r,4).Text) | $($wsDp.Cells.Item($r,5).Text)"
}
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

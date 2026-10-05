$xl = New-Object -ComObject Excel.Application
$wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
$ws = $wb.Sheets.Item("CAPACITY_PLAN")
Write-Host "ListObjects on CAPACITY_PLAN: $($ws.ListObjects.Count)"
foreach ($lo in $ws.ListObjects) { Write-Host "  LO: $($lo.Name) Range: $($lo.Range.Address())" }
$lr = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
Write-Host "Rows on CAPACITY_PLAN: $lr"
for ($c = 1; $c -le 9; $c++) {
    Write-Host "Col ${c}: $($ws.Cells.Item(1, $c).Value2)"
}
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

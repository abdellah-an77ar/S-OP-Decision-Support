$xl = New-Object -ComObject Excel.Application
$wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
foreach ($sn in @("SALES_HISTORY", "FORECAST", "DEMAND_PLAN", "CAPACITY_PLAN", "INVENTORY_PLAN", "LOG")) {
    $ws = $wb.Sheets.Item($sn)
    $lr = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    Write-Host "$sn Rows: $lr"
}
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

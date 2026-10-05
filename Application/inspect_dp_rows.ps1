$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $wsDp = $wb.Sheets.Item("DEMAND_PLAN")
    $dpLast = $wsDp.Cells.Item($wsDp.Rows.Count, 1).End(-4162).Row
    Write-Host "dpLast = $dpLast"
    for ($r = 2; $r -le [Math]::Min(10, $dpLast); $r++) {
        $v1 = $wsDp.Cells.Item($r, 1).Value2
        $v2 = $wsDp.Cells.Item($r, 2).Value2
        $v5 = $wsDp.Cells.Item($r, 5).Value2
        Write-Host "Row ${r}: SKU=$v1, Date=$v2, Qty=$v5"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Calling ImporterDonnees..."
    $xl.Run("AtlasFood_SOP.xlsm!ImporterDonnees")
    Write-Host "Finished ImporterDonnees"

    $wsLog = $wb.Sheets.Item("LOG")
    $lr = $wsLog.Cells.Item($wsLog.Rows.Count, 1).End(-4162).Row
    Write-Host "Log rows: $lr"
    for ($r = 2; $r -le $lr; $r++) {
        $t = $wsLog.Cells.Item($r, 1).Text
        $m = $wsLog.Cells.Item($r, 3).Text
        $lvl = $wsLog.Cells.Item($r, 4).Text
        $msg = $wsLog.Cells.Item($r, 5).Text
        Write-Host "$t | $m | $lvl | $msg"
    }

    $sh = $wb.Sheets.Item("SALES_HISTORY")
    $slr = $sh.Cells.Item($sh.Rows.Count, 1).End(-4162).Row
    Write-Host "SALES_HISTORY rows: $slr"
} catch {
    Write-Host "Caught error: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

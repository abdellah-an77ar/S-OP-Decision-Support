$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $v = $xl.Run("AtlasFood_SOP.xlsm!GetParam", "SEUIL_ALERTE_GAP")
    Write-Host "GetParam SEUIL_ALERTE_GAP = $v"
    $c = $xl.Run("AtlasFood_SOP.xlsm!GetCycleID")
    Write-Host "GetCycleID = $c"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

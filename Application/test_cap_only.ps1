$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Workbook opened. Calling LancerCapacityPlan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    Write-Host "LancerCapacityPlan returned!"
    $wb.Save()
    Write-Host "Saved!"
} catch {
    Write-Host "Caught error: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

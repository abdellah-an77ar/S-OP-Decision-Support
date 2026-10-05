$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "1. Running ImporterDonnees..."
    $xl.Run("AtlasFood_SOP.xlsm!ImporterDonnees")
    $wb.Save()
    Write-Host "   Import done and saved. SALES_HISTORY rows: $($wb.Sheets.Item('SALES_HISTORY').Cells.Item($wb.Sheets.Item('SALES_HISTORY').Rows.Count, 1).End(-4162).Row)"

    Write-Host "2. Running LancerForecast..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerForecast")
    $wb.Save()
    Write-Host "   Forecast done and saved. FORECAST rows: $($wb.Sheets.Item('FORECAST').Cells.Item($wb.Sheets.Item('FORECAST').Rows.Count, 1).End(-4162).Row)"

    Write-Host "3. Running LancerDemandPlan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerDemandPlan")
    $wb.Save()
    Write-Host "   Demand Plan done and saved. DEMAND_PLAN rows: $($wb.Sheets.Item('DEMAND_PLAN').Cells.Item($wb.Sheets.Item('DEMAND_PLAN').Rows.Count, 1).End(-4162).Row)"

    Write-Host "4. Running LancerCapacityPlan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    $wb.Save()
    Write-Host "   Capacity Plan done and saved. CAPACITY_PLAN rows: $($wb.Sheets.Item('CAPACITY_PLAN').Cells.Item($wb.Sheets.Item('CAPACITY_PLAN').Rows.Count, 1).End(-4162).Row)"

} catch {
    Write-Host "ERROR: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

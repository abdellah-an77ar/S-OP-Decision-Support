$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Running Capacity Plan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    Write-Host "Capacity Plan Done. Rows in CAPACITY_PLAN: " $wb.Sheets("CAPACITY_PLAN").Cells(100, 1).End(-4162).Row

    Write-Host "Running Inventory Plan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerInventoryPlan")
    Write-Host "Inventory Plan Done. Rows in INVENTORY_PLAN: " $wb.Sheets("INVENTORY_PLAN").Cells(100, 1).End(-4162).Row

    Write-Host "Running Gap Analysis..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerGapAnalysis")
    Write-Host "Gap Analysis Done. Rows in GAP_ANALYSIS: " $wb.Sheets("GAP_ANALYSIS").Cells(100, 1).End(-4162).Row

    Write-Host "Running Scenarios..."
    $xl.Run("AtlasFood_SOP.xlsm!GenererScenarios")
    Write-Host "Scenarios Done. Rows in SCENARIOS: " $wb.Sheets("SCENARIOS").Cells(100, 1).End(-4162).Row

    Write-Host "Running CalculerKPI..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $xl.Run("AtlasFood_SOP.xlsm!CalculerKPI")
    $sw.Stop()
    Write-Host "CalculerKPI Done in $($sw.ElapsedMilliseconds) ms. Rows in KPI_RESULTS: " $wb.Sheets("KPI_RESULTS").Cells(100, 1).End(-4162).Row

    $wb.Save()
    Write-Host "Saved successfully!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

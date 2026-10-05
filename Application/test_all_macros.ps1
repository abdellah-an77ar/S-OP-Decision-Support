$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Workbook opened."

    function Test-Macro([string]$name, [scriptblock]$sb) {
        Write-Host "Testing $name..." -NoNewline
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            & $sb
            $sw.Stop()
            Write-Host " [PASS] ($($sw.ElapsedMilliseconds) ms)" -ForegroundColor Green
        } catch {
            $sw.Stop()
            Write-Host " [FAIL] ($($sw.ElapsedMilliseconds) ms)" -ForegroundColor Red
            Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Test-Macro "ImporterDonnees" { $xl.Run("AtlasFood_SOP.xlsm!ImporterDonnees") }
    Test-Macro "LancerForecast" { $xl.Run("AtlasFood_SOP.xlsm!LancerForecast") }
    Test-Macro "LancerDemandPlan" { $xl.Run("AtlasFood_SOP.xlsm!LancerDemandPlan") }
    Test-Macro "LancerCapacityPlan" { $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan") }
    Test-Macro "LancerInventoryPlan" { $xl.Run("AtlasFood_SOP.xlsm!LancerInventoryPlan") }
    Test-Macro "LancerGapAnalysis" { $xl.Run("AtlasFood_SOP.xlsm!LancerGapAnalysis") }
    Test-Macro "GenererScenarios" { $xl.Run("AtlasFood_SOP.xlsm!GenererScenarios") }
    Test-Macro "CalculerKPI" { $xl.Run("AtlasFood_SOP.xlsm!CalculerKPI") }
    Test-Macro "ValiderPlan" { $xl.Run("AtlasFood_SOP.xlsm!ValiderPlan", "SC-04") }
    Test-Macro "LancerMonteCarlo" { $xl.Run("AtlasFood_SOP.xlsm!LancerMonteCarlo") }
    Test-Macro "ExporterVersPowerBI" { $xl.Run("AtlasFood_SOP.xlsm!ExporterVersPowerBI") }

    $wb.Save()
    Write-Host "`nAll macros executed and workbook saved!" -ForegroundColor Cyan
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

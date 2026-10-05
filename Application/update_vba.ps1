$WorkspaceRoot = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
$xlsmPath = Join-Path $WorkspaceRoot "Application\AtlasFood_SOP.xlsm"
$basDir   = Join-Path $WorkspaceRoot "VBA_Modules"

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $wb = $excel.Workbooks.Open($xlsmPath)
    Write-Host "Workbook opened: $xlsmPath"
    $vbaProj = $wb.VBProject

    $modNames = @("mod01_CONFIG", "mod02_IMPORT", "mod03_FORECAST", "mod04_PLANS", "mod05_SCENARIOS", "mod06_KPI", "mod07_EXPORT")
    foreach ($m in $modNames) {
        $basFile = Join-Path $basDir "$m.bas"
        if (Test-Path $basFile) {
            foreach ($c in $vbaProj.VBComponents) {
                if ($c.Name -eq $m) {
                    $vbaProj.VBComponents.Remove($c)
                    break
                }
            }
            $vbc = $vbaProj.VBComponents.Import($basFile)
            Write-Host "Imported $m ($($vbc.CodeModule.CountOfLines) lines)"
        }
    }

    $twPath = Join-Path $basDir "ThisWorkbook.bas"
    if (Test-Path $twPath) {
        $twCode = Get-Content $twPath -Raw -Encoding UTF8
        $twComp = $vbaProj.VBComponents.Item("ThisWorkbook")
        if ($twComp.CodeModule.CountOfLines -gt 0) { $twComp.CodeModule.DeleteLines(1, $twComp.CodeModule.CountOfLines) }
        $twComp.CodeModule.InsertLines(1, $twCode)
        Write-Host "Updated ThisWorkbook"
    }

    $wb.Save()
    Write-Host "Workbook saved successfully with updated VBA!"
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $ws = $wb.Sheets.Item("PRODUCT")
    Write-Host "PRODUCT ListObjects before: $($ws.ListObjects.Count)"
    foreach ($lo in $ws.ListObjects) { Write-Host "  LO: $($lo.Name)" }

    # Let's test reading CSV in PowerShell to see line counts
    $csvPath = "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Data\PRODUCT.csv"
    $lines = Get-Content $csvPath
    Write-Host "CSV Lines in PRODUCT.csv: $($lines.Count)"
    Write-Host "Line 0: $($lines[0])"
    Write-Host "Line 1: $($lines[1])"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

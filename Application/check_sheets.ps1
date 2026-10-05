$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "=== Sheets Status in AtlasFood_SOP.xlsm ==="
    foreach ($sh in $wb.Sheets) {
        $lr = $sh.Cells.Item($sh.Rows.Count, 1).End(-4162).Row
        $loCount = $sh.ListObjects.Count
        Write-Host "$($sh.Name.PadRight(20)) : Rows=$lr, Tables=$loCount"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

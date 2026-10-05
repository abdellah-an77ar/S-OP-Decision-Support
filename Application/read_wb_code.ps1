$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $comp = $wb.VBProject.VBComponents.Item("mod04_PLANS")
    $count = $comp.CodeModule.CountOfLines
    Write-Host "Total lines in mod04_PLANS: $count"
    $lines = $comp.CodeModule.Lines(1, 85)
    Write-Host $lines
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

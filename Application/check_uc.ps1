$xl = New-Object -ComObject Excel.Application
Write-Host "UserControl initial:" $xl.UserControl
$wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
Write-Host "UserControl after Open:" $xl.UserControl
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

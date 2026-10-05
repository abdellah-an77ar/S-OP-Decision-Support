$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "Workbook opened."
    # Let's trigger compilation via VBE
    $vbe = $xl.VBE
    Write-Host "VBE ActiveVBProject: $($vbe.ActiveVBProject.Name)"
    $compileCmd = $xl.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
    if ($compileCmd -ne $null) {
        Write-Host "Compile command found: Enabled=$($compileCmd.Enabled)"
        if ($compileCmd.Enabled) {
            Write-Host "Executing compile..."
            $compileCmd.Execute()
            Write-Host "Compile executed!"
        } else {
            Write-Host "Project already compiled!"
        }
    } else {
        Write-Host "Compile command not found."
    }
} catch {
    Write-Host "Error: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

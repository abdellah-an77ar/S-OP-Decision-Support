$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $comp = $wb.VBProject.VBComponents.Add(1) # std module
    $comp.CodeModule.AddFromString(@"
Sub DirectTest()
    On Error GoTo EH
    Debug.Print "DirectTest calling LancerCapacityPlan..."
    Call LancerCapacityPlan
    Debug.Print "DirectTest returned!"
    Exit Sub
EH:
    Debug.Print "DirectTest error: " & Err.Description
End Sub
"@)
    Write-Host "Calling DirectTest..."
    $xl.Run("DirectTest")
    Write-Host "DirectTest completed successfully!"
} catch {
    Write-Host "Caught: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

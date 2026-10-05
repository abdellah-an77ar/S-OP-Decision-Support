$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.CodeModule.AddFromString(@"
Public Sub TestDbg()
    Dim fn As Integer: fn = FreeFile
    Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_cap.txt" For Append As #fn
    Print #fn, "Test from TestDbg"
    Close #fn
End Sub
"@)
    Write-Host "Calling LancerCapacityPlan..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    Write-Host "Returned from LancerCapacityPlan!"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

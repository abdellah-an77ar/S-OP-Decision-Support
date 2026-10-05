$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.CodeModule.AddFromString(@"
Sub TestDict()
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add "A", 1
    d.Add "B", 2
    Dim k As Variant
    For Each k In d.Keys()
        ' ok
    Next k
End Sub
"@)
    $xl.Run("TestDict")
    Write-Host "TestDict passed!"
} catch {
    Write-Host "TestDict ERROR: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

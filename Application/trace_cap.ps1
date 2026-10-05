$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    Write-Host "1. Workbook opened"
    $wsDp = $wb.Sheets.Item("DEMAND_PLAN")
    $wsCap = $wb.Sheets.Item("CAPACITY_PLAN")
    $wsPlant = $wb.Sheets.Item("PLANT")
    $wsProd = $wb.Sheets.Item("PRODUCT")

    Write-Host "2. Reading tbl_PLANT"
    $loPlant = $wsPlant.ListObjects.Item("tbl_PLANT")
    Write-Host "   tbl_PLANT rows: $($loPlant.DataBodyRange.Rows.Count)"

    Write-Host "3. Reading tbl_PRODUCT"
    $loProd = $wsProd.ListObjects.Item("tbl_PRODUCT")
    Write-Host "   tbl_PRODUCT rows: $($loProd.DataBodyRange.Rows.Count)"

    Write-Host "4. Reading DEMAND_PLAN"
    $dpLast = $wsDp.Cells.Item($wsDp.Rows.Count, 1).End(-4162).Row
    Write-Host "   DEMAND_PLAN rows: $dpLast"

    for ($dr = 2; $dr -le [Math]::Min(5, $dpLast); $dr++) {
        $sku = $wsDp.Cells.Item($dr, 1).Value2
        $dt = $wsDp.Cells.Item($dr, 2).Value2
        $q = $wsDp.Cells.Item($dr, 5).Value2
        Write-Host "   Row ${dr}: SKU=$sku, Date=$dt, Qty=$q"
    }

    Write-Host "5. Now calling VBA LancerCapacityPlan directly via COM..."
    $xl.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan")
    Write-Host "6. Returned successfully from LancerCapacityPlan!"
    $capLast = $wsCap.Cells.Item($wsCap.Rows.Count, 1).End(-4162).Row
    Write-Host "   CAPACITY_PLAN rows: $capLast"
} catch {
    Write-Host "ERROR: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

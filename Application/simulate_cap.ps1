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

    Write-Host "2. Testing ClearSheet via VBA macro ClearSheet..."
    $xl.Run("AtlasFood_SOP.xlsm!ClearSheet", "CAPACITY_PLAN")
    Write-Host "   ClearSheet returned successfully!"

    Write-Host "3. Testing reading tbl_PLANT..."
    $loPlant = $wsPlant.ListObjects.Item("tbl_PLANT")
    $dictCap = @{}
    for ($pr = 1; $pr -le $loPlant.DataBodyRange.Rows.Count; $pr++) {
        $plantId = [string]$loPlant.DataBodyRange.Item($pr, 1).Value2
        $cRef = [double]$loPlant.DataBodyRange.Item($pr, 3).Value2
        $mg = [double]$loPlant.DataBodyRange.Item($pr, 4).Value2
        $hsPct = [double]$loPlant.DataBodyRange.Item($pr, 5).Value2
        $stAbs = [double]$loPlant.DataBodyRange.Item($pr, 6).Value2
        $dictCap[$plantId] = "$($cRef * (1 + $mg))|$($cRef * $hsPct)|$stAbs"
        Write-Host "   Plant $plantId -> $($dictCap[$plantId])"
    }

    Write-Host "4. Testing reading tbl_PRODUCT..."
    $loProd = $wsProd.ListObjects.Item("tbl_PRODUCT")
    $spMap = @{}
    for ($pp = 1; $pp -le $loProd.DataBodyRange.Rows.Count; $pp++) {
        $sid = [string]$loProd.DataBodyRange.Item($pp, 1).Value2
        $pld = [string]$loProd.DataBodyRange.Item($pp, 4).Value2
        $spMap[$sid] = $pld
        Write-Host "   SKU $sid -> Plant $pld"
    }

    Write-Host "5. Testing aggregating charge from DEMAND_PLAN..."
    $dpLast = $wsDp.Cells.Item($wsDp.Rows.Count, 1).End(-4162).Row
    $dictCharge = @{}
    for ($dr = 2; $dr -le $dpLast; $dr++) {
        $dSku = [string]$wsDp.Cells.Item($dr, 1).Value2
        $dDate = [DateTime]::FromOADate($wsDp.Cells.Item($dr, 2).Value2)
        $dQty = [double]$wsDp.Cells.Item($dr, 5).Value2
        $dPlant = if ($spMap.ContainsKey($dSku)) { $spMap[$dSku] } else { "INCONNU" }
        $ck = "$dPlant|$($dDate.ToString('yyyy-MM'))"
        if ($dictCharge.ContainsKey($ck)) { $dictCharge[$ck] += $dQty } else { $dictCharge[$ck] = $dQty }
    }
    Write-Host "   Aggregated $($dictCharge.Count) plant/month buckets."

    Write-Host "6. Writing capacity plan rows to CAPACITY_PLAN sheet..."
    $rowOut = 2
    foreach ($k in $dictCharge.Keys) {
        $parts = $k.Split("|")
        $kPlant = $parts[0]
        $kMois = [DateTime]::ParseExact($parts[1] + "-01", "yyyy-MM-dd", $null)
        $charge = $dictCharge[$k]
        $capParts = if ($dictCap.ContainsKey($kPlant)) { $dictCap[$kPlant].Split("|") } else { @("5000","1000","500") }
        $capN = [double]$capParts[0]
        $capHS = [double]$capParts[1]
        $capST = [double]$capParts[2]
        $util = if ($capN -gt 0) { $charge / $capN } else { 0 }
        $st = if ($util -gt 1.0) { "SURCHARGE" } elseif ($util -gt 0.85) { "ALERTE" } else { "OK" }

        $wsCap.Cells.Item($rowOut, 1).Value2 = $kPlant
        $wsCap.Cells.Item($rowOut, 2).Value2 = $kMois.ToString("yyyy-MM-01")
        $wsCap.Cells.Item($rowOut, 3).Value2 = [Math]::Round($capN, 0)
        $wsCap.Cells.Item($rowOut, 4).Value2 = [Math]::Round($capHS, 0)
        $wsCap.Cells.Item($rowOut, 5).Value2 = [Math]::Round($capST, 0)
        $wsCap.Cells.Item($rowOut, 6).Value2 = [Math]::Round($charge, 0)
        $wsCap.Cells.Item($rowOut, 7).Value2 = [Math]::Round($util * 100, 1)
        $wsCap.Cells.Item($rowOut, 8).Value2 = $st
        $wsCap.Cells.Item($rowOut, 9).Value2 = "CYC-2023-06"
        $rowOut++
    }
    Write-Host "   Written $($rowOut - 2) rows."

    Write-Host "7. Creating tbl_CAPACITY ListObject..."
    $rng = $wsCap.Range("A1:I" + ($rowOut - 1))
    $newLo = $wsCap.ListObjects.Add(1, $rng, [System.Reflection.Missing]::Value, 1)
    $newLo.Name = "tbl_CAPACITY"
    $newLo.TableStyle = "TableStyleMedium5"
    Write-Host "   ListObject created successfully: $($newLo.Name)"

    $wb.Save()
    Write-Host "Simulation completed successfully and saved!"
} catch {
    Write-Host "ERROR at Step: $_ at line $($_.InvocationInfo.ScriptLineNumber)"
    Write-Host $_.ScriptStackTrace
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

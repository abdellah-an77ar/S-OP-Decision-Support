# ================================================================
# SCRIPT CORRECTIF v2 - Ajoute les boutons et leur liaison VBA
# ================================================================
param(
    [string]$XlsmPath = "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm"
)
$ErrorActionPreference = "Stop"

Write-Host "Ouverture du fichier..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open($XlsmPath)
$wsAcc = $wb.Sheets["ACCUEIL"]

# Supprimer anciens boutons
foreach ($ole in @($wsAcc.OLEObjects)) {
    try { $ole.Delete() } catch {}
}

Write-Host "Ajout des boutons..." -ForegroundColor Yellow

# Structure: Caption, Left, Top, Width, Height, BackColor (entier long VBA compatible)
$btnDefs = @(
    "Importer les donnees|14|220|205|42|3046326",
    "Lancer le Forecast|230|220|180|42|7372099",
    "Demand Plan|422|220|160|42|7372099",
    "Capacity Inventory Gap|594|220|230|42|16763904",
    "Generer les Scenarios|14|275|200|42|7352480",
    "Calculer les KPI|226|275|170|42|7352480",
    "Valider le Plan Final|408|275|190|42|43571",
    "Monte Carlo Expert|610|275|165|42|12583168",
    "CYCLE COMPLET|14|340|190|42|15570995",
    "Exporter vers Power BI|216|340|195|42|5921370",
    "Reinitialiser les calculs|423|340|185|42|11885081",
    "Voir le LOG|620|340|135|42|5921370"
)

$created = @()
foreach ($def in $btnDefs) {
    $parts = $def -split "\|"
    $cap  = $parts[0]
    $left = [double]$parts[1]
    $top  = [double]$parts[2]
    $w    = [double]$parts[3]
    $h    = [double]$parts[4]
    $bgc  = [int]$parts[5]
    try {
        $ole = $wsAcc.OLEObjects.Add("Forms.CommandButton.1", $null, $false, $false, $null, $null, $null, $left, $top, $w, $h)
        $ole.Object.Caption   = $cap
        $ole.Object.BackColor = $bgc
        $ole.Object.ForeColor = 16777215  # Blanc
        $ole.Object.Font.Bold = $true
        $ole.Object.Font.Size = 8
        $safe = ($cap -replace "[^A-Za-z0-9]","_") -replace "_+","_"
        $safe = $safe.Trim("_")
        $oleName = "btn_$safe"
        $ole.Name = $oleName
        $created += $oleName
        Write-Host "   OK: $oleName" -ForegroundColor Green
    } catch {
        Write-Host "   WARN: $cap - $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

# Injecter le code VBA de liaison depuis un fichier temporaire
$accCode = @'
Option Explicit

Private Sub btn_Importer_les_donnees_Click()
    Call ImporterDonnees
End Sub

Private Sub btn_Lancer_le_Forecast_Click()
    Call LancerForecast
End Sub

Private Sub btn_Demand_Plan_Click()
    Call LancerDemandPlan
End Sub

Private Sub btn_Capacity_Inventory_Gap_Click()
    Call LancerCapacityPlan
    Call LancerInventoryPlan
    Call LancerGapAnalysis
End Sub

Private Sub btn_Generer_les_Scenarios_Click()
    Call GenererScenarios
End Sub

Private Sub btn_Calculer_les_KPI_Click()
    Call CalculerKPI
End Sub

Private Sub btn_Valider_le_Plan_Final_Click()
    Call ValiderPlanFinal
End Sub

Private Sub btn_Monte_Carlo_Expert_Click()
    Call LancerMonteCarlo
End Sub

Private Sub btn_CYCLE_COMPLET_Click()
    Call LancerCycleComplet
End Sub

Private Sub btn_Exporter_vers_Power_BI_Click()
    Call ExporterVersPowerBI
End Sub

Private Sub btn_Reinitialiser_les_calculs_Click()
    Call ReinitialiserCalculs
End Sub

Private Sub btn_Voir_le_LOG_Click()
    ThisWorkbook.Sheets("LOG").Visible = -1
    ThisWorkbook.Sheets("LOG").Activate
End Sub
'@

$vbaProj = $wb.VBProject
foreach ($vbc in $vbaProj.VBComponents) {
    try {
        $shName = $vbc.Properties.Item("Name").Value
        if ($shName -eq "ACCUEIL") {
            if ($vbc.CodeModule.CountOfLines -gt 0) {
                $vbc.CodeModule.DeleteLines(1, $vbc.CodeModule.CountOfLines)
            }
            $vbc.CodeModule.InsertLines(1, $accCode)
            Write-Host "Code VBA liaison injecte dans Sheet ACCUEIL" -ForegroundColor Green
            break
        }
    } catch {}
}

Write-Host "Sauvegarde..." -ForegroundColor Yellow
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  BOUTONS AJOUTES - Fichier mis a jour" -ForegroundColor Cyan
Write-Host "  $XlsmPath" -ForegroundColor Cyan
Write-Host "  Boutons crees: $($created.Count)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

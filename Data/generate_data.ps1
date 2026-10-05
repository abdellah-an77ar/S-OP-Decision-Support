# =============================================================
# AtlasFood S&OP - Generateur de donnees synthetiques realistes
# Source: Inspire du dataset SupplyGraph (arXiv 2401.15299)
# Nature: SYNTHETIQUE - Test uniquement
# Date: 2026-09-29
# =============================================================

Add-Type -AssemblyName System.Globalization

$culture = [System.Globalization.CultureInfo]::InvariantCulture

# Chemin de sortie
$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# -------------------------------------------------------
# PRODUITS (sous-ensemble de 5 SKU sur 3 usines)
# -------------------------------------------------------
$products = @(
    [pscustomobject]@{SKU_ID="SKU-01"; Nom="Biscuits Nature 250g";  Famille="Biscuits"; Plant_ID="USN-01"; Poids_unit=0.25; DLC_jours=180; Stock_secu_methode="SS_NORMAL"}
    [pscustomobject]@{SKU_ID="SKU-02"; Nom="Biscuits Chocolat 250g";Famille="Biscuits"; Plant_ID="USN-01"; Poids_unit=0.25; DLC_jours=150; Stock_secu_methode="SS_NORMAL"}
    [pscustomobject]@{SKU_ID="SKU-03"; Nom="Crackers Sale 200g";    Famille="Crackers"; Plant_ID="USN-02"; Poids_unit=0.20; DLC_jours=120; Stock_secu_methode="SS_NORMAL"}
    [pscustomobject]@{SKU_ID="SKU-04"; Nom="Crackers Cereales 200g";Famille="Crackers"; Plant_ID="USN-02"; Poids_unit=0.20; DLC_jours=120; Stock_secu_methode="SS_NORMAL"}
    [pscustomobject]@{SKU_ID="SKU-05"; Nom="Gaufrettes Vanille 150g";Famille="Gaufrettes";Plant_ID="USN-03";Poids_unit=0.15; DLC_jours=90; Stock_secu_methode="SS_SEASONAL"}
)
$products | Export-Csv "$outDir\PRODUCT.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "PRODUCT.csv cree."

# -------------------------------------------------------
# USINES
# -------------------------------------------------------
$plants = @(
    [pscustomobject]@{Plant_ID="USN-01"; Nom="Usine Nord";   Capacite_ref=5000; Marge_cap=0.10; Cap_HS_max_pct=0.20; Cap_ST_max=1000}
    [pscustomobject]@{Plant_ID="USN-02"; Nom="Usine Sud";    Capacite_ref=4000; Marge_cap=0.10; Cap_HS_max_pct=0.20; Cap_ST_max=800}
    [pscustomobject]@{Plant_ID="USN-03"; Nom="Usine Est";    Capacite_ref=2500; Marge_cap=0.10; Cap_HS_max_pct=0.15; Cap_ST_max=500}
)
$plants | Export-Csv "$outDir\PLANT.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "PLANT.csv cree."

# -------------------------------------------------------
# CALENDRIER (Jan-Aout 2023)
# -------------------------------------------------------
$calendar = @()
$months = @(
    @{Mois="2023-01-01"; Jours_ouvres=22; Jours_feries=0; Jours_Ramadan=0}
    @{Mois="2023-02-01"; Jours_ouvres=20; Jours_feries=0; Jours_Ramadan=0}
    @{Mois="2023-03-01"; Jours_ouvres=23; Jours_feries=1; Jours_Ramadan=22}  # Ramadan debut mars
    @{Mois="2023-04-01"; Jours_ouvres=18; Jours_feries=2; Jours_Ramadan=8}   # Fin Ramadan + Eid
    @{Mois="2023-05-01"; Jours_ouvres=22; Jours_feries=1; Jours_Ramadan=0}
    @{Mois="2023-06-01"; Jours_ouvres=21; Jours_feries=1; Jours_Ramadan=0}
    @{Mois="2023-07-01"; Jours_ouvres=21; Jours_feries=0; Jours_Ramadan=0}
    @{Mois="2023-08-01"; Jours_ouvres=7;  Jours_feries=0; Jours_Ramadan=0}   # Partiel (9 jours)
)
foreach ($m in $months) {
    $calendar += [pscustomobject]$m
}
$calendar | Export-Csv "$outDir\CALENDAR.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "CALENDAR.csv cree."

# -------------------------------------------------------
# COUTS UNITAIRES
# -------------------------------------------------------
$costs = @(
    [pscustomobject]@{Type="HS";         Plant_ID="USN-01"; Montant_unit=350; Unite_monet="UM"}
    [pscustomobject]@{Type="HS";         Plant_ID="USN-02"; Montant_unit=340; Unite_monet="UM"}
    [pscustomobject]@{Type="HS";         Plant_ID="USN-03"; Montant_unit=320; Unite_monet="UM"}
    [pscustomobject]@{Type="ST";         Plant_ID="USN-01"; Montant_unit=500; Unite_monet="UM"}
    [pscustomobject]@{Type="ST";         Plant_ID="USN-02"; Montant_unit=490; Unite_monet="UM"}
    [pscustomobject]@{Type="ST";         Plant_ID="USN-03"; Montant_unit=480; Unite_monet="UM"}
    [pscustomobject]@{Type="Stockage";   Plant_ID="USN-01"; Montant_unit=12;  Unite_monet="UM"}
    [pscustomobject]@{Type="Stockage";   Plant_ID="USN-02"; Montant_unit=12;  Unite_monet="UM"}
    [pscustomobject]@{Type="Stockage";   Plant_ID="USN-03"; Montant_unit=11;  Unite_monet="UM"}
    [pscustomobject]@{Type="Penurie";    Plant_ID="USN-01"; Montant_unit=800; Unite_monet="UM"}
    [pscustomobject]@{Type="Penurie";    Plant_ID="USN-02"; Montant_unit=800; Unite_monet="UM"}
    [pscustomobject]@{Type="Penurie";    Plant_ID="USN-03"; Montant_unit=750; Unite_monet="UM"}
    [pscustomobject]@{Type="Peremption"; Plant_ID="USN-01"; Montant_unit=200; Unite_monet="UM"}
    [pscustomobject]@{Type="Peremption"; Plant_ID="USN-02"; Montant_unit=200; Unite_monet="UM"}
    [pscustomobject]@{Type="Peremption"; Plant_ID="USN-03"; Montant_unit=190; Unite_monet="UM"}
    [pscustomobject]@{Type="Prod";       Plant_ID="USN-01"; Montant_unit=150; Unite_monet="UM"}
    [pscustomobject]@{Type="Prod";       Plant_ID="USN-02"; Montant_unit=145; Unite_monet="UM"}
    [pscustomobject]@{Type="Prod";       Plant_ID="USN-03"; Montant_unit=140; Unite_monet="UM"}
)
$costs | Export-Csv "$outDir\COST.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "COST.csv cree."

# -------------------------------------------------------
# HISTORIQUE QUOTIDIEN (Jan 1 -> Aout 9, 2023)
# Donnees synthetiques realistes (inspire SupplyGraph FMCG Bangladesh)
# -------------------------------------------------------
$startDate = [datetime]"2023-01-01"
$endDate   = [datetime]"2023-08-09"
$random    = [System.Random]::new(42)  # seed fixe pour reproductibilite

# Profils de demande par SKU
$profiles = @{
    "SKU-01" = @{base=250; trend=0.5; seasonMarch=1.3; seasonApril=1.5; cv=0.12; prodFactor=1.05}
    "SKU-02" = @{base=200; trend=0.3; seasonMarch=1.2; seasonApril=1.4; cv=0.15; prodFactor=1.05}
    "SKU-03" = @{base=180; trend=0.2; seasonMarch=1.1; seasonApril=1.2; cv=0.18; prodFactor=1.08}
    "SKU-04" = @{base=150; trend=0.1; seasonMarch=1.1; seasonApril=1.15; cv=0.20; prodFactor=1.10}
    "SKU-05" = @{base=120; trend=0.4; seasonMarch=1.5; seasonApril=1.8; cv=0.25; prodFactor=1.12}
}

$salesHistory    = [System.Collections.ArrayList]::new()
$prodHistory     = [System.Collections.ArrayList]::new()
$deliveryHistory = [System.Collections.ArrayList]::new()

$currentDate = $startDate
$dayIndex = 0
while ($currentDate -le $endDate) {
    $dayOfWeek = $currentDate.DayOfWeek
    $isWorkday = ($dayOfWeek -ne [DayOfWeek]::Friday -and $dayOfWeek -ne [DayOfWeek]::Saturday)  # Bangladesh: vendredi+samedi repos
    
    foreach ($sku in $profiles.Keys) {
        $p = $profiles[$sku]
        $prod = ($products | Where-Object { $_.SKU_ID -eq $sku })[0]
        
        if ($isWorkday) {
            # Facteur saisonnier
            $seasonFactor = 1.0
            if ($currentDate.Month -eq 3) { $seasonFactor = $p.seasonMarch }
            if ($currentDate.Month -eq 4) { $seasonFactor = $p.seasonApril }
            
            # Tendance
            $trendFactor = 1.0 + ($p.trend / 100) * $dayIndex
            
            # Bruit gaussien (Box-Muller)
            $u1 = 1.0 - $random.NextDouble()
            $u2 = 1.0 - $random.NextDouble()
            $noise = [Math]::Sqrt(-2.0 * [Math]::Log($u1)) * [Math]::Cos(2.0 * [Math]::PI * $u2)
            
            $baseDemand = $p.base * $seasonFactor * $trendFactor
            $demandRaw  = [Math]::Max(0, $baseDemand + $noise * $baseDemand * $p.cv)
            $demand     = [Math]::Round($demandRaw)
            
            # Production = demande * facteur prod (avec variation aleatoire)
            $prodRaw    = $demand * $p.prodFactor * (1 + $noise * 0.05)
            $production = [Math]::Max(0, [Math]::Round($prodRaw))
            
            # Livraisons = 90-95% de la production (delai de 1-2 jours)
            $delivFactor = 0.90 + $random.NextDouble() * 0.05
            $delivery    = [Math]::Round($production * $delivFactor)
            
            $dateStr = $currentDate.ToString("yyyy-MM-dd", $culture)
            
            $null = $salesHistory.Add([pscustomobject]@{
                SKU_ID         = $sku
                Date           = $dateStr
                Qte_commandee  = $demand
                Unite          = "Unite"
            })
            
            $null = $prodHistory.Add([pscustomobject]@{
                SKU_ID      = $sku
                Plant_ID    = $prod.Plant_ID
                Date        = $dateStr
                Qte_produite = $production
            })
            
            $null = $deliveryHistory.Add([pscustomobject]@{
                SKU_ID     = $sku
                Date       = $dateStr
                Qte_livree = $delivery
            })
        }
    }
    
    $currentDate = $currentDate.AddDays(1)
    $dayIndex++
}

$salesHistory    | Export-Csv "$outDir\SALES_HISTORY.csv"    -NoTypeInformation -Encoding UTF8 -Delimiter ";"
$prodHistory     | Export-Csv "$outDir\PRODUCTION_HISTORY.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"
$deliveryHistory | Export-Csv "$outDir\DELIVERY_HISTORY.csv"  -NoTypeInformation -Encoding UTF8 -Delimiter ";"

Write-Host "SALES_HISTORY.csv cree: $($salesHistory.Count) lignes."
Write-Host "PRODUCTION_HISTORY.csv cree: $($prodHistory.Count) lignes."
Write-Host "DELIVERY_HISTORY.csv cree: $($deliveryHistory.Count) lignes."
Write-Host ""
Write-Host "=== Generation des donnees terminee ==="
Write-Host "Nature: SYNTHETIQUE - inspiree de SupplyGraph (arXiv 2401.15299)"
Write-Host "Reproductible: seed=42"

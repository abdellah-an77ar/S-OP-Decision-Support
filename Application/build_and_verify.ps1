# ================================================================
# BUILD AND VERIFY - AtlasFood S&OP DSS
# Reconstruit entierement AtlasFood_SOP.xlsm, injecte le VBA,
# cree les boutons Shapes OnAction, et execute le test E2E complet
# ================================================================
param(
    [string]$WorkspaceRoot = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
)
$ErrorActionPreference = "Stop"
$appDir   = Join-Path $WorkspaceRoot "Application"
$xlsmPath = Join-Path $appDir "AtlasFood_SOP.xlsm"
$basDir   = Join-Path $WorkspaceRoot "VBA_Modules"
$dataDir  = Join-Path $WorkspaceRoot "Data"
$pbiDir   = Join-Path $WorkspaceRoot "PowerBI"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AtlasFood S&OP DSS - Construction & Audit Complet" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Arreter tout processus Excel
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800

if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

Write-Host "[1/9] Demarrage Excel COM..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1 # Low security for programmatic build

$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets[$wb.Sheets.Count].Delete() }
$wb.Sheets[1].Name = "ACCUEIL"

# ================================================================
Write-Host "[2/9] Creation des 20 feuilles..." -ForegroundColor Yellow
$shDefs = @(
    "ACCUEIL|1F4975", "PARAMETRES|70AD47", "PLANT|4472C4", "PRODUCT|4472C4",
    "CALENDAR|4472C4", "COST|4472C4", "SALES_HISTORY|ED7D31",
    "PRODUCTION_HISTORY|ED7D31", "DELIVERY_HISTORY|ED7D31",
    "FORECAST|FFC000", "DEMAND_PLAN|FFC000", "CAPACITY_PLAN|FFC000",
    "INVENTORY_PLAN|FFC000", "GAP_ANALYSIS|FF0000", "SCENARIOS|7030A0",
    "PLAN_FINAL|00B050", "KPI_RESULTS|00B050", "ALERTES|FF0000",
    "EXPORT_POWERBI|595959", "LOG|595959"
)
foreach ($sd in $shDefs) {
    $parts = $sd -split "\|"
    $name = $parts[0]; $hex = $parts[1]
    $ws = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq $name) { $ws = $s; break } }
    if ($null -eq $ws) {
        $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets[$wb.Sheets.Count])
        $ws.Name = $name
    }
    $r  = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g  = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b2 = [Convert]::ToInt32($hex.Substring(4,2),16)
    $ws.Tab.Color = $r + ($g -shl 8) + ($b2 -shl 16)
}

function SetHeader {
    param($ws, [string[]]$hdrs, [int]$bgCol = 0x1F4975)
    for ($c = 0; $c -lt $hdrs.Count; $c++) {
        $ws.Cells.Item(1, $c+1).Value2 = $hdrs[$c]
    }
    $lc = [char](64 + $hdrs.Count)
    $rng = $ws.Range("A1:${lc}1")
    $rng.Font.Bold = $true
    $rng.Interior.Color = $bgCol
    $rng.Font.Color = 0xFFFFFF
    $rng.RowHeight = 24
    $rng.VerticalAlignment = -4108
}

function MkTable {
    param($ws, [string]$nm, [int]$lr, [int]$lc, [string]$sty = "TableStyleMedium2")
    $ec = [char](64 + $lc)
    $rng = $ws.Range("A1:${ec}${lr}")
    $lo = $ws.ListObjects.Add(1, $rng, [System.Reflection.Missing]::Value, 1)
    $lo.Name = $nm
    $lo.TableStyle = $sty
}

# ================================================================
Write-Host "[3/9] Donnees de reference et structures..." -ForegroundColor Yellow

# PLANT
$wsP = $wb.Sheets["PLANT"]
SetHeader $wsP @("Plant_ID","Nom","Capacite_ref","Marge_cap","Cap_HS_max_pct","Cap_ST_max")
$wsP.Cells.Item(2,1).Value2="USN-01"; $wsP.Cells.Item(2,2).Value2="Usine Nord";  $wsP.Cells.Item(2,3).Value2=5000; $wsP.Cells.Item(2,4).Value2=0.10; $wsP.Cells.Item(2,5).Value2=0.20; $wsP.Cells.Item(2,6).Value2=1000
$wsP.Cells.Item(3,1).Value2="USN-02"; $wsP.Cells.Item(3,2).Value2="Usine Sud";   $wsP.Cells.Item(3,3).Value2=4000; $wsP.Cells.Item(3,4).Value2=0.10; $wsP.Cells.Item(3,5).Value2=0.20; $wsP.Cells.Item(3,6).Value2=800
$wsP.Cells.Item(4,1).Value2="USN-03"; $wsP.Cells.Item(4,2).Value2="Usine Est";   $wsP.Cells.Item(4,3).Value2=2500; $wsP.Cells.Item(4,4).Value2=0.10; $wsP.Cells.Item(4,5).Value2=0.15; $wsP.Cells.Item(4,6).Value2=500
$wsP.Columns("A:F").AutoFit()
MkTable $wsP "tbl_PLANT" 4 6

# PRODUCT
$wsPr = $wb.Sheets["PRODUCT"]
SetHeader $wsPr @("SKU_ID","Nom","Famille","Plant_ID","Poids_unit","DLC_jours","Stock_secu_methode")
$wsPr.Cells.Item(2,1).Value2="SKU-01"; $wsPr.Cells.Item(2,2).Value2="Biscuits Nature 250g";    $wsPr.Cells.Item(2,3).Value2="Biscuits";   $wsPr.Cells.Item(2,4).Value2="USN-01"; $wsPr.Cells.Item(2,5).Value2=0.25; $wsPr.Cells.Item(2,6).Value2=180; $wsPr.Cells.Item(2,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(3,1).Value2="SKU-02"; $wsPr.Cells.Item(3,2).Value2="Biscuits Chocolat 250g";  $wsPr.Cells.Item(3,3).Value2="Biscuits";   $wsPr.Cells.Item(3,4).Value2="USN-01"; $wsPr.Cells.Item(3,5).Value2=0.25; $wsPr.Cells.Item(3,6).Value2=150; $wsPr.Cells.Item(3,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(4,1).Value2="SKU-03"; $wsPr.Cells.Item(4,2).Value2="Crackers Sale 200g";      $wsPr.Cells.Item(4,3).Value2="Crackers";   $wsPr.Cells.Item(4,4).Value2="USN-02"; $wsPr.Cells.Item(4,5).Value2=0.20; $wsPr.Cells.Item(4,6).Value2=120; $wsPr.Cells.Item(4,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(5,1).Value2="SKU-04"; $wsPr.Cells.Item(5,2).Value2="Crackers Cereales 200g";  $wsPr.Cells.Item(5,3).Value2="Crackers";   $wsPr.Cells.Item(5,4).Value2="USN-02"; $wsPr.Cells.Item(5,5).Value2=0.20; $wsPr.Cells.Item(5,6).Value2=120; $wsPr.Cells.Item(5,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(6,1).Value2="SKU-05"; $wsPr.Cells.Item(6,2).Value2="Gaufrettes Vanille 150g"; $wsPr.Cells.Item(6,3).Value2="Gaufrettes"; $wsPr.Cells.Item(6,4).Value2="USN-03"; $wsPr.Cells.Item(6,5).Value2=0.15; $wsPr.Cells.Item(6,6).Value2=90;  $wsPr.Cells.Item(6,7).Value2="SS_SEASONAL"
$wsPr.Columns("A:G").AutoFit()
MkTable $wsPr "tbl_PRODUCT" 6 7

# CALENDAR
$wsCal = $wb.Sheets["CALENDAR"]
SetHeader $wsCal @("Mois","Jours_ouvres","Jours_feries","Jours_Ramadan")
$wsCal.Cells.Item(2,1).Value2="2023-01-01"; $wsCal.Cells.Item(2,2).Value2=22; $wsCal.Cells.Item(2,3).Value2=0;  $wsCal.Cells.Item(2,4).Value2=0
$wsCal.Cells.Item(3,1).Value2="2023-02-01"; $wsCal.Cells.Item(3,2).Value2=20; $wsCal.Cells.Item(3,3).Value2=0;  $wsCal.Cells.Item(3,4).Value2=0
$wsCal.Cells.Item(4,1).Value2="2023-03-01"; $wsCal.Cells.Item(4,2).Value2=23; $wsCal.Cells.Item(4,3).Value2=1;  $wsCal.Cells.Item(4,4).Value2=22
$wsCal.Cells.Item(5,1).Value2="2023-04-01"; $wsCal.Cells.Item(5,2).Value2=18; $wsCal.Cells.Item(5,3).Value2=2;  $wsCal.Cells.Item(5,4).Value2=8
$wsCal.Cells.Item(6,1).Value2="2023-05-01"; $wsCal.Cells.Item(6,2).Value2=22; $wsCal.Cells.Item(6,3).Value2=1;  $wsCal.Cells.Item(6,4).Value2=0
$wsCal.Cells.Item(7,1).Value2="2023-06-01"; $wsCal.Cells.Item(7,2).Value2=21; $wsCal.Cells.Item(7,3).Value2=1;  $wsCal.Cells.Item(7,4).Value2=0
$wsCal.Cells.Item(8,1).Value2="2023-07-01"; $wsCal.Cells.Item(8,2).Value2=21; $wsCal.Cells.Item(8,3).Value2=0;  $wsCal.Cells.Item(8,4).Value2=0
$wsCal.Cells.Item(9,1).Value2="2023-08-01"; $wsCal.Cells.Item(9,2).Value2=7;  $wsCal.Cells.Item(9,3).Value2=0;  $wsCal.Cells.Item(9,4).Value2=0
$wsCal.Columns("A:D").AutoFit()
MkTable $wsCal "tbl_CALENDAR" 9 4

# COST
$wsCst = $wb.Sheets["COST"]
SetHeader $wsCst @("Type","Plant_ID","Montant_unit","Unite_monet")
$costRows = @(
    "HS|USN-01|350|UM","HS|USN-02|340|UM","HS|USN-03|320|UM",
    "ST|USN-01|500|UM","ST|USN-02|490|UM","ST|USN-03|480|UM",
    "Stockage|USN-01|12|UM","Stockage|USN-02|12|UM","Stockage|USN-03|11|UM",
    "Penurie|USN-01|800|UM","Penurie|USN-02|800|UM","Penurie|USN-03|750|UM",
    "Peremption|USN-01|200|UM","Peremption|USN-02|200|UM","Peremption|USN-03|190|UM",
    "Prod|USN-01|150|UM","Prod|USN-02|145|UM","Prod|USN-03|140|UM"
)
for ($ri = 0; $ri -lt $costRows.Count; $ri++) {
    $cp = $costRows[$ri] -split "\|"
    $wsCst.Cells.Item($ri+2,1).Value2 = $cp[0]
    $wsCst.Cells.Item($ri+2,2).Value2 = $cp[1]
    $wsCst.Cells.Item($ri+2,3).Value2 = [double]$cp[2]
    $wsCst.Cells.Item($ri+2,4).Value2 = $cp[3]
}
$wsCst.Columns("A:D").AutoFit()
MkTable $wsCst "tbl_COST" ($costRows.Count+1) 4

# PARAMETRES
$wsPm = $wb.Sheets["PARAMETRES"]
SetHeader $wsPm @("Cle","Valeur","Unite","Nature","Description") 0x70AD47
$pmRows = @(
    "HORIZON_MOIS|6|mois|Parametre|Horizon de planification S.O.P.",
    "NIVEAU_SERVICE|0.95|pct|Parametre|Niveau de service cible (95%)",
    "DELAI_LIVRAISON|5|jours|Parametre|Delai reapprovisionnement moyen",
    "MARGE_CAPACITE|0.10|pct|Hypothese|Marge sur capacite observee (10%)",
    "SEUIL_ALERTE_GAP|0.85|pct|Parametre|Seuil alerte utilisation capacite (85%)",
    "SEUIL_SURSTOCK|2|x|Parametre|Ratio stock/demande declenchant surstock (2x)",
    "CV_SEUIL_STABLE|0.20||Parametre|CV max classe STABLE",
    "CV_SEUIL_VOLATILE|0.50||Parametre|CV min classe VOLATILE",
    "BACKTEST_SEMAINES|8|sem|Parametre|Semaines pour le back-test forecast",
    "DATE_COUPURE|2023-05-31|date|Parametre|Date de coupure T0",
    "CYCLE_ID|CYC-2023-06||Parametre|Identifiant cycle courant",
    "UNITE_CHARGE|Tonne||Parametre|Unite de charge",
    "DEVISE|UM||Hypothese|Unite monetaire neutre",
    "STOCK_INIT_SKU01|5000|unites|Hypothese|Stock initial SKU-01",
    "STOCK_INIT_SKU02|4000|unites|Hypothese|Stock initial SKU-02",
    "STOCK_INIT_SKU03|3500|unites|Hypothese|Stock initial SKU-03",
    "STOCK_INIT_SKU04|3000|unites|Hypothese|Stock initial SKU-04",
    "STOCK_INIT_SKU05|2000|unites|Hypothese|Stock initial SKU-05",
    "MC_ITERATIONS|1000||Parametre|Iterations Monte Carlo",
    "MC_SIGMA_PCT|0.15|pct|Hypothese|Ecart-type relatif Monte Carlo",
    "SEUIL_VALIDATEUR|SOP Manager||Parametre|Profil validateur plan"
)
for ($ri = 0; $ri -lt $pmRows.Count; $ri++) {
    $pp = $pmRows[$ri] -split "\|"
    $wsPm.Cells.Item($ri+2,1).Value2 = $pp[0]
    $wsPm.Cells.Item($ri+2,2).Value2 = $pp[1]
    $wsPm.Cells.Item($ri+2,3).Value2 = $pp[2]
    $wsPm.Cells.Item($ri+2,4).Value2 = $pp[3]
    $wsPm.Cells.Item($ri+2,5).Value2 = $pp[4]
    if ($pp[3] -eq "Hypothese") { $wsPm.Cells.Item($ri+2,4).Font.Italic = $true }
}
$wsPm.Columns("A").ColumnWidth = 24; $wsPm.Columns("B").ColumnWidth = 16; $wsPm.Columns("E").ColumnWidth = 48
MkTable $wsPm "tbl_PARAM" ($pmRows.Count+1) 5

# En-tetes historiques
$histInfo = @{
    "SALES_HISTORY"      = "SKU_ID|Date|Qte_commandee|Unite"
    "PRODUCTION_HISTORY" = "SKU_ID|Plant_ID|Date|Qte_produite"
    "DELIVERY_HISTORY"   = "SKU_ID|Date|Qte_livree"
}
foreach ($hn in $histInfo.Keys) {
    $wsH = $wb.Sheets[$hn]
    $hh  = $histInfo[$hn] -split "\|"
    SetHeader $wsH $hh 0xED7D31
}

# En-tetes feuilles calculees
$calcInfo = @{
    "FORECAST"       = "SKU_ID|Periode|Qte_prevue|Methode|WAPE_pct|Biais_pct|Statut|Cycle_ID|Classe|bgCol=1F4975"
    "DEMAND_PLAN"    = "SKU_ID|Mois|Qte_stat|Ajustement|Qte_finale|Cycle_ID|Commentaire|bgCol=1F4975"
    "CAPACITY_PLAN"  = "Plant_ID|Mois|Capacite_normale|Cap_HS_max|Cap_ST_max|Charge|Utilisation_pct|Statut|Cycle_ID|bgCol=1F4975"
    "INVENTORY_PLAN" = "SKU_ID|Mois|Stock_debut|Stock_secu|Production|Livraisons|Stock_fin|DLC_restante_est|Statut|Cycle_ID|bgCol=1F4975"
    "GAP_ANALYSIS"   = "Plant_ID|Mois|Besoin_net|Cap_disponible|Gap|Gap_pct|Statut|Cycle_ID|bgCol=FF0000"
    "SCENARIOS"      = "Scenario_ID|Nom|Hypotheses|Cout_total|Taux_service_pct|Stock_final_tot|Risque_peremption|Util_max_pct|Cycle_ID|Statut|bgCol=7030A0"
    "PLAN_FINAL"     = "Plan_ID|Version|SKU_ID|Plant_ID|Mois|Prod_normale|Prod_HS|Prod_ST|Stock_fin|Demande|Non_servi|Scenario_retenu|Statut|Valideur|Horodatage|bgCol=00B050"
    "KPI_RESULTS"    = "KPI_ID|Indicateur|Perimetre|Periode|Valeur|Unite|Seuil_vert|Seuil_rouge|Statut_seuil|Cycle_ID|bgCol=00B050"
    "ALERTES"        = "Alert_ID|Type|Criticite|Perimetre|Message|Cycle_ID|Horodatage|bgCol=FF0000"
    "EXPORT_POWERBI" = "Table|Derniere_export|Nb_lignes|Statut|bgCol=595959"
    "LOG"            = "Horodatage|Utilisateur|Module|Niveau|Message|bgCol=595959"
}
foreach ($cn in $calcInfo.Keys) {
    $wsC = $wb.Sheets[$cn]
    $rawCols = $calcInfo[$cn] -split "\|"
    $bgHex = "1F4975"
    $hdrs  = @()
    foreach ($col in $rawCols) {
        if ($col -match "^bgCol=(.+)$") { $bgHex = $matches[1] }
        else { $hdrs += $col }
    }
    $bgI = [Convert]::ToInt32($bgHex.Substring(0,2),16) + ([Convert]::ToInt32($bgHex.Substring(2,2),16) -shl 8) + ([Convert]::ToInt32($bgHex.Substring(4,2),16) -shl 16)
    SetHeader $wsC $hdrs $bgI
    $wsC.Columns("A:P").ColumnWidth = 16
}

# ================================================================
Write-Host "[4/9] Mise en page ecran ACCUEIL..." -ForegroundColor Yellow
$wsAcc = $wb.Sheets["ACCUEIL"]
$wsAcc.Cells.Interior.Color = 0xF8FAFC

# Colonne A marge
$wsAcc.Columns("A:A").ColumnWidth = 3

# Titre Bandeau
$wsAcc.Range("B2:M4").Merge()
$wsAcc.Cells.Item(2,2).Value2 = "AtlasFood S" + [char]38 + "OP Decision Support System"
$wsAcc.Cells.Item(2,2).Font.Size = 20
$wsAcc.Cells.Item(2,2).Font.Bold = $true
$wsAcc.Cells.Item(2,2).Font.Color = 0xFFFFFF
$wsAcc.Cells.Item(2,2).HorizontalAlignment = -4108
$wsAcc.Cells.Item(2,2).VerticalAlignment   = -4108
$wsAcc.Range("B2:M4").Interior.Color = 0x1E3A8A # Navy Blue

$wsAcc.Range("B5:M5").Merge()
$wsAcc.Cells.Item(5,2).Value2 = "Application d'aide a la decision S&OP - Cycle mensuel integre (VBA + Power BI)"
$wsAcc.Cells.Item(5,2).Font.Size = 11
$wsAcc.Cells.Item(5,2).Font.Italic = $true
$wsAcc.Cells.Item(5,2).Font.Color = 0x1E3A8A
$wsAcc.Cells.Item(5,2).HorizontalAlignment = -4108
$wsAcc.Rows.Item(5).RowHeight = 22

$wsAcc.Range("B6:M6").Interior.Color = 0x10B981 # Emerald accent line
$wsAcc.Rows.Item(6).RowHeight = 4

# Bloc Metadonnees
$wsAcc.Rows.Item(8).RowHeight = 24
$wsAcc.Cells.Item(8,2).Value2 = "Utilisateur :"; $wsAcc.Cells.Item(8,2).Font.Bold = $true
$wsAcc.Cells.Item(8,3).Value2 = $env:USERNAME;    $wsAcc.Cells.Item(8,3).Interior.Color = 0xFFFFFF; $wsAcc.Cells.Item(8,3).Borders.LineStyle = 1
$wsAcc.Cells.Item(8,5).Value2 = "Cycle S&OP :";  $wsAcc.Cells.Item(8,5).Font.Bold = $true
$wsAcc.Cells.Item(8,6).Value2 = "CYC-2023-06";    $wsAcc.Cells.Item(8,6).Interior.Color = 0xFFFFFF; $wsAcc.Cells.Item(8,6).Borders.LineStyle = 1
$wsAcc.Cells.Item(8,8).Value2 = "Date Coupure :";$wsAcc.Cells.Item(8,8).Font.Bold = $true
$wsAcc.Cells.Item(8,9).Value2 = "31/05/2023";     $wsAcc.Cells.Item(8,9).Interior.Color = 0xFFFFFF; $wsAcc.Cells.Item(8,9).Borders.LineStyle = 1
$wsAcc.Cells.Item(8,11).Value2= "Version :";     $wsAcc.Cells.Item(8,11).Font.Bold = $true
$wsAcc.Cells.Item(8,12).Value2= "1.0 - Production";$wsAcc.Cells.Item(8,12).Font.Color = 0x059669

# Workflow banner
$wsAcc.Range("B10:M10").Merge()
$wsAcc.Cells.Item(10,2).Value2 = "WORKFLOW DU CYCLE S&OP (Cliquez sur les etapes de gauche a droite)"
$wsAcc.Cells.Item(10,2).Font.Bold = $true; $wsAcc.Cells.Item(10,2).Font.Size = 11; $wsAcc.Cells.Item(10,2).Font.Color = 0x1E3A8A
$wsAcc.Range("B10:M10").Interior.Color = 0xDBEAFE
$wsAcc.Rows.Item(10).RowHeight = 24

# Labels etapes
$stepInfo = @(
    @("J1 : DONNEES", 2, 0x1D4ED8),
    @("J2 : DEMANDE", 5, 0x059669),
    @("J3 : SUPPLY",  8, 0xD97706),
    @("J4 : SCENARIOS", 11, 0x7C3AED)
)
foreach ($st in $stepInfo) {
    $c = $st[1]
    $wsAcc.Cells.Item(12, $c).Value2 = $st[0]
    $wsAcc.Cells.Item(12, $c).Font.Bold = $true
    $wsAcc.Cells.Item(12, $c).Font.Color = 0xFFFFFF
    $wsAcc.Cells.Item(12, $c).Interior.Color = $st[2]
    $wsAcc.Cells.Item(12, $c).HorizontalAlignment = -4108
}
$wsAcc.Rows.Item(12).RowHeight = 20

# Zone Statut dynamique
$wsAcc.Cells.Item(26,2).Value2 = "STATUT DU PROCESSUS :"
$wsAcc.Cells.Item(26,2).Font.Bold = $true; $wsAcc.Cells.Item(26,2).Font.Color = 0x1E3A8A
$wsAcc.Range("B27:M27").Merge()
$wsAcc.Cells.Item(27,2).Value2 = "Application initialisee. Pret pour l'etape J1 (Importation des donnees)."
$wsAcc.Cells.Item(27,2).Font.Italic = $true
$wsAcc.Cells.Item(27,2).Font.Color = 0x1E3A8A
$wsAcc.Range("B27:M27").Interior.Color = 0xFEF3C7 # Light amber
$wsAcc.Range("B27:M27").Borders.LineStyle = 1
$wsAcc.Rows.Item(27).RowHeight = 24

# Source des donnees
$wsAcc.Cells.Item(29,2).Value2 = "Provenance des donnees :"
$wsAcc.Cells.Item(29,2).Font.Bold = $true
$wsAcc.Range("C29:M29").Merge()
$wsAcc.Cells.Item(29,3).Value2 = "Donnees synthetiques inspirees du benchmark SupplyGraph (arXiv:2401.15299) | Dossier /Data"
$wsAcc.Cells.Item(29,3).Font.Italic = $true
$wsAcc.Cells.Item(29,3).Font.Color = 0x64748B

# Noms definis
$wsAcc.Cells.Item(8,3).Name  = "B_USER_NAME"
$wsAcc.Cells.Item(27,2).Name = "B_STATUS"

# ================================================================
Write-Host "[5/9] Ajout des 12 boutons avec OnAction..." -ForegroundColor Yellow

$btnSpecs = @(
    # Ligne 1 : Workflow J1 a J3
    @{Text="1. Importer les donnees (J1)";   Macro="ImporterDonnees";    Left=30;  Top=230; Width=190; Height=36; Color=0x2563EB},
    @{Text="2. Lancer Forecast (J2)";        Macro="LancerForecast";     Left=235; Top=230; Width=180; Height=36; Color=0x10B981},
    @{Text="3. Demand Plan (J2b)";           Macro="LancerDemandPlan";   Left=430; Top=230; Width=170; Height=36; Color=0x10B981},
    @{Text="4. Capacity - Stock - Gap (J3)"; Macro="LancerSupplyPlan";   Left=615; Top=230; Width=210; Height=36; Color=0xD97706},

    # Ligne 2 : Workflow J4 a J5 + Expert
    @{Text="5. Generer Scenarios (J4)";      Macro="GenererScenarios";   Left=30;  Top=280; Width=190; Height=36; Color=0x7C3AED},
    @{Text="6. Calculer les KPI (J4b)";      Macro="CalculerKPI";        Left=235; Top=280; Width=180; Height=36; Color=0x7C3AED},
    @{Text="7. Valider Plan Final (J5)";     Macro="ValiderPlanFinal";   Left=430; Top=280; Width=170; Height=36; Color=0x059669},
    @{Text="8. Monte Carlo Expert";          Macro="LancerMonteCarlo";   Left=615; Top=280; Width=210; Height=36; Color=0xDC2626},

    # Ligne 3 : Automatisation & Outils
    @{Text="CYCLE COMPLET (J2-J4)";          Macro="LancerCycleComplet"; Left=30;  Top=340; Width=190; Height=34; Color=0xEA580C},
    @{Text="Exporter vers Power BI";         Macro="ExporterVersPowerBI";Left=235; Top=340; Width=180; Height=34; Color=0x475569},
    @{Text="Reinitialiser les calculs";      Macro="ReinitialiserCalculs";Left=430;Top=340; Width=170; Height=34; Color=0x9A3412},
    @{Text="Journal LOG";                    Macro="AfficherLog";        Left=615; Top=340; Width=210; Height=34; Color=0x475569}
)

foreach ($b in $btnSpecs) {
    try {
        # msoShapeRoundedRectangle = 5
        $shp = $wsAcc.Shapes.AddShape(5, $b.Left, $b.Top, $b.Width, $b.Height)
        $shp.Name = "btn_" + ($b.Macro)
        $shp.TextFrame.Characters().Text = $b.Text
        $shp.TextFrame.Characters().Font.Bold = $true
        $shp.TextFrame.Characters().Font.Size = 9
        $shp.TextFrame.Characters().Font.Color = 0xFFFFFF
        $shp.TextFrame.HorizontalAlignment = -4108 # Center
        $shp.TextFrame.VerticalAlignment   = -4108 # Center
        $shp.Fill.Solid()
        $shp.Fill.ForeColor.RGB = $b.Color
        $shp.Line.Visible = $false
        $shp.OnAction = $b.Macro
        Write-Host "   -> Bouton '$($b.Text)' relié à macro '$($b.Macro)'" -ForegroundColor Green
    } catch {
        Write-Host "   [WARN] Erreur bouton $($b.Text): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ================================================================
Write-Host "[6/9] Injection des modules VBA via VBComponents.Import..." -ForegroundColor Yellow
$vbaProj = $wb.VBProject

$modNames = @("mod01_CONFIG", "mod02_IMPORT", "mod03_FORECAST", "mod04_PLANS", "mod05_SCENARIOS", "mod06_KPI", "mod07_EXPORT")
foreach ($m in $modNames) {
    $basFile = Join-Path $basDir "$m.bas"
    if (-not (Test-Path $basFile)) { Write-Host "   [WARN] Manquant: $basFile" -ForegroundColor Red; continue }
    
    # Supprimer composant si deja existant
    foreach ($c in $vbaProj.VBComponents) {
        if ($c.Name -eq $m) {
            $vbaProj.VBComponents.Remove($c)
            break
        }
    }
    $vbc = $vbaProj.VBComponents.Import($basFile)
    Write-Host "   -> Module $m importe avec succes ($($vbc.CodeModule.CountOfLines) lignes)" -ForegroundColor Gray
}

# ThisWorkbook
$twPath = Join-Path $basDir "ThisWorkbook.bas"
if (Test-Path $twPath) {
    $twCode = Get-Content $twPath -Raw -Encoding UTF8
    $twComp = $vbaProj.VBComponents["ThisWorkbook"]
    if ($twComp.CodeModule.CountOfLines -gt 0) { $twComp.CodeModule.DeleteLines(1, $twComp.CodeModule.CountOfLines) }
    $twComp.CodeModule.InsertLines(1, $twCode)
    Write-Host "   -> ThisWorkbook configure ($($twComp.CodeModule.CountOfLines) lignes)" -ForegroundColor Gray
}

# ================================================================
Write-Host "[7/9] Sauvegarde initiale en .xlsm..." -ForegroundColor Yellow
$wb.SaveAs($xlsmPath, 52)
Write-Host "   OK: $xlsmPath sauvegarde avec succes !" -ForegroundColor Green

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep -Milliseconds 1000

# ================================================================
Write-Host "[8/9] REOUVERTURE ET TEST D'EXECUTION COMPLET (E2E)..." -ForegroundColor Yellow

$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
$excel2.EnableEvents = $false
$excel2.AutomationSecurity = 1 # msoAutomationSecurityLow

$wb2 = $excel2.Workbooks.Open($xlsmPath)
Write-Host "   Fichier $xlsmPath ouvert avec macros activees (AutomationSecurity=1)" -ForegroundColor Green

$testResults = @()

function Run-Step {
    param([string]$name, [scriptblock]$action)
    Write-Host "   Execution: $name..." -NoNewline
    try {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        & $action
        $sw.Stop()
        Write-Host " [PASS] ($($sw.ElapsedMilliseconds) ms)" -ForegroundColor Green
        return @{Step=$name; Status="PASS"; Error=$null; Ms=$sw.ElapsedMilliseconds}
    } catch {
        Write-Host " [FAIL]" -ForegroundColor Red
        Write-Host "      Erreur: $($_.Exception.Message)" -ForegroundColor Red
        return @{Step=$name; Status="FAIL"; Error=$_.Exception.Message; Ms=0}
    }
}

# 1. Import donnees
$res1 = Run-Step "1. ImporterDonnees" { $excel2.Run("AtlasFood_SOP.xlsm!ImporterDonnees") }
$testResults += $res1

# 2. Forecast
$res2 = Run-Step "2. LancerForecast" { $excel2.Run("AtlasFood_SOP.xlsm!LancerForecast") }
$testResults += $res2

# 3. Demand Plan
$res3 = Run-Step "3. LancerDemandPlan" { $excel2.Run("AtlasFood_SOP.xlsm!LancerDemandPlan") }
$testResults += $res3

# 4. Capacity Plan
$res4 = Run-Step "4. LancerCapacityPlan" { $excel2.Run("AtlasFood_SOP.xlsm!LancerCapacityPlan") }
$testResults += $res4

# 5. Inventory Plan
$res5 = Run-Step "5. LancerInventoryPlan" { $excel2.Run("AtlasFood_SOP.xlsm!LancerInventoryPlan") }
$testResults += $res5

# 6. Gap Analysis
$res6 = Run-Step "6. LancerGapAnalysis" { $excel2.Run("AtlasFood_SOP.xlsm!LancerGapAnalysis") }
$testResults += $res6

# 7. Scenarios
$res7 = Run-Step "7. GenererScenarios" { $excel2.Run("AtlasFood_SOP.xlsm!GenererScenarios") }
$testResults += $res7

# 8. KPI
$res8 = Run-Step "8. CalculerKPI" { $excel2.Run("AtlasFood_SOP.xlsm!CalculerKPI") }
$testResults += $res8

# 9. Validation Plan Final
$res9 = Run-Step "9. ValiderPlan (SC-04)" { $excel2.Run("AtlasFood_SOP.xlsm!ValiderPlan", "SC-04") }
$testResults += $res9

# 10. Monte Carlo Expert
$res10 = Run-Step "10. LancerMonteCarlo" { $excel2.Run("AtlasFood_SOP.xlsm!LancerMonteCarlo") }
$testResults += $res10

# 11. Export Power BI
$res11 = Run-Step "11. ExporterVersPowerBI" { $excel2.Run("AtlasFood_SOP.xlsm!ExporterVersPowerBI") }
$testResults += $res11

# ================================================================
Write-Host "[9/9] Verification approfondie des resultats generes..." -ForegroundColor Yellow

$sheetChecks = @(
    "SALES_HISTORY", "PRODUCTION_HISTORY", "DELIVERY_HISTORY",
    "FORECAST", "DEMAND_PLAN", "CAPACITY_PLAN", "INVENTORY_PLAN",
    "GAP_ANALYSIS", "SCENARIOS", "PLAN_FINAL", "KPI_RESULTS", "ALERTES", "LOG"
)

Write-Host "`n   --- Synthese des lignes par feuille Excel ---" -ForegroundColor Cyan
foreach ($sn in $sheetChecks) {
    $sh = $wb2.Sheets[$sn]
    $lr = $sh.Cells.Item($sh.Rows.Count, 1).End(-4162).Row # xlUp
    $dataRows = [math]::Max(0, $lr - 1)
    $statusColor = if ($dataRows -gt 0) { "Green" } else { "Red" }
    Write-Host "   Feuille: $($sn.PadRight(22)) -> $dataRows lignes de donnees" -ForegroundColor $statusColor
}

# Verifier les fichiers CSV exportes dans /PowerBI
Write-Host "`n   --- Fichiers CSV exportes dans /PowerBI ---" -ForegroundColor Cyan
$pbiFiles = Get-ChildItem $pbiDir -Filter "PBI_*.csv" -ErrorAction SilentlyContinue
foreach ($pf in $pbiFiles) {
    $lc = (Get-Content $pf.FullName | Measure-Object -Line).Lines
    Write-Host "   CSV: $($pf.Name.PadRight(24)) -> $lc lignes ($([math]::Round($pf.Length/1024, 1)) Ko)" -ForegroundColor Green
}

# Sauvegarder classeur avec les donnees calculees
$wb2.Save()

# Fermeture propre
$wb2.Close($false)
$excel2.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AUDIT TECHNIQUE ET EXECUTION E2E TERMINES" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan


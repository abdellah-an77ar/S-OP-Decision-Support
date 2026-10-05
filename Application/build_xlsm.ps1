# ================================================================
# BUILD SCRIPT v3 - AtlasFood S&OP DSS
# Construction du fichier Excel .xlsm via COM
# Les modules VBA sont lus depuis les fichiers .bas
# ================================================================
param(
    [string]$WorkspaceRoot = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
)
$ErrorActionPreference = "Stop"
$appDir   = Join-Path $WorkspaceRoot "Application"
$xlsmPath = Join-Path $appDir "AtlasFood_SOP.xlsm"
$basDir   = Join-Path $WorkspaceRoot "VBA_Modules"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AtlasFood S&OP DSS - Construction XLSM" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

Write-Host "[1/8] Demarrage Excel..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets[$wb.Sheets.Count].Delete() }
$wb.Sheets[1].Name = "ACCUEIL"

# ================================================================
Write-Host "[2/8] Creation des 20 feuilles..." -ForegroundColor Yellow
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
Write-Host "   OK: $($wb.Sheets.Count) feuilles" -ForegroundColor Green

# ================================================================
Write-Host "[3/8] Donnees de reference..." -ForegroundColor Yellow

function SetHeader {
    param($ws, [string[]]$hdrs, [int]$bgCol = 0x1F4975)
    for ($c = 0; $c -lt $hdrs.Count; $c++) {
        $ws.Cells.Item(1, $c+1).Value2 = $hdrs[$c]
    }
    $lc = [char](64 + $hdrs.Count)
    $rng = $ws.Range("A1:${lc}1")
    $rng.Font.Bold   = $true
    $rng.Interior.Color = $bgCol
    $rng.Font.Color  = [int]0xFFFFFF
}
function MkTable {
    param($ws, [string]$nm, [int]$lr, [int]$lc, [string]$sty = "TableStyleMedium2")
    $ec = [char](64 + $lc)
    $rng = $ws.Range("A1:${ec}${lr}")
    $lo = $ws.ListObjects.Add(1, $rng, [System.Reflection.Missing]::Value, 1)
    $lo.Name = $nm; $lo.TableStyle = $sty
}

# PLANT
$wsP = $wb.Sheets["PLANT"]
SetHeader $wsP @("Plant_ID","Nom","Capacite_ref","Marge_cap","Cap_HS_max_pct","Cap_ST_max")
$wsP.Cells.Item(2,1).Value2="USN-01"; $wsP.Cells.Item(2,2).Value2="Usine Nord";  $wsP.Cells.Item(2,3).Value2=5000; $wsP.Cells.Item(2,4).Value2=0.10; $wsP.Cells.Item(2,5).Value2=0.20; $wsP.Cells.Item(2,6).Value2=1000
$wsP.Cells.Item(3,1).Value2="USN-02"; $wsP.Cells.Item(3,2).Value2="Usine Sud";   $wsP.Cells.Item(3,3).Value2=4000; $wsP.Cells.Item(3,4).Value2=0.10; $wsP.Cells.Item(3,5).Value2=0.20; $wsP.Cells.Item(3,6).Value2=800
$wsP.Cells.Item(4,1).Value2="USN-03"; $wsP.Cells.Item(4,2).Value2="Usine Est";   $wsP.Cells.Item(4,3).Value2=2500; $wsP.Cells.Item(4,4).Value2=0.10; $wsP.Cells.Item(4,5).Value2=0.15; $wsP.Cells.Item(4,6).Value2=500
MkTable $wsP "tbl_PLANT" 4 6

# PRODUCT
$wsPr = $wb.Sheets["PRODUCT"]
SetHeader $wsPr @("SKU_ID","Nom","Famille","Plant_ID","Poids_unit","DLC_jours","Stock_secu_methode")
$wsPr.Cells.Item(2,1).Value2="SKU-01"; $wsPr.Cells.Item(2,2).Value2="Biscuits Nature 250g";    $wsPr.Cells.Item(2,3).Value2="Biscuits";   $wsPr.Cells.Item(2,4).Value2="USN-01"; $wsPr.Cells.Item(2,5).Value2=0.25; $wsPr.Cells.Item(2,6).Value2=180; $wsPr.Cells.Item(2,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(3,1).Value2="SKU-02"; $wsPr.Cells.Item(3,2).Value2="Biscuits Chocolat 250g";  $wsPr.Cells.Item(3,3).Value2="Biscuits";   $wsPr.Cells.Item(3,4).Value2="USN-01"; $wsPr.Cells.Item(3,5).Value2=0.25; $wsPr.Cells.Item(3,6).Value2=150; $wsPr.Cells.Item(3,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(4,1).Value2="SKU-03"; $wsPr.Cells.Item(4,2).Value2="Crackers Sale 200g";      $wsPr.Cells.Item(4,3).Value2="Crackers";   $wsPr.Cells.Item(4,4).Value2="USN-02"; $wsPr.Cells.Item(4,5).Value2=0.20; $wsPr.Cells.Item(4,6).Value2=120; $wsPr.Cells.Item(4,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(5,1).Value2="SKU-04"; $wsPr.Cells.Item(5,2).Value2="Crackers Cereales 200g";  $wsPr.Cells.Item(5,3).Value2="Crackers";   $wsPr.Cells.Item(5,4).Value2="USN-02"; $wsPr.Cells.Item(5,5).Value2=0.20; $wsPr.Cells.Item(5,6).Value2=120; $wsPr.Cells.Item(5,7).Value2="SS_NORMAL"
$wsPr.Cells.Item(6,1).Value2="SKU-05"; $wsPr.Cells.Item(6,2).Value2="Gaufrettes Vanille 150g"; $wsPr.Cells.Item(6,3).Value2="Gaufrettes"; $wsPr.Cells.Item(6,4).Value2="USN-03"; $wsPr.Cells.Item(6,5).Value2=0.15; $wsPr.Cells.Item(6,6).Value2=90;  $wsPr.Cells.Item(6,7).Value2="SS_SEASONAL"
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
MkTable $wsCst "tbl_COST" ($costRows.Count+1) 4

# PARAMETRES
$wsPm = $wb.Sheets["PARAMETRES"]
SetHeader $wsPm @("Cle","Valeur","Unite","Nature","Description") 0x70AD47
$pmRows = @(
    "HORIZON_MOIS|6|mois|Parametre|Horizon de planification S.O.P.",
    "NIVEAU_SERVICE|0.95|pct|Parametre|Niveau de service cible",
    "DELAI_LIVRAISON|5|jours|Parametre|Delai reapprovisionnement moyen",
    "MARGE_CAPACITE|0.10|pct|Hypothese|Marge sur capacite observee",
    "SEUIL_ALERTE_GAP|0.85|pct|Parametre|Seuil alerte utilisation capacite",
    "SEUIL_SURSTOCK|2|x|Parametre|Ratio stock/demande declenchant surstock",
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
$wsPm.Columns("A").ColumnWidth = 22; $wsPm.Columns("B").ColumnWidth = 16; $wsPm.Columns("E").ColumnWidth = 48
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
    $wsH.Cells.Item(2,1).Value2 = "(vide - cliquez Importer les donnees)"
    $wsH.Cells.Item(2,1).Font.Italic = $true
    $wsH.Cells.Item(2,1).Font.Color  = [int]0x808080
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
Write-Host "   OK: donnees de reference" -ForegroundColor Green

# ================================================================
Write-Host "[4/8] Interface ACCUEIL..." -ForegroundColor Yellow
$wsAcc = $wb.Sheets["ACCUEIL"]
$wsAcc.Cells.Interior.Color = [int]0xF2F2F2

# Masquer colonne A
$wsAcc.Columns("A:A").Hidden = $true

# Titre
$wsAcc.Range("B2:N4").Merge()
$ampStr = "AtlasFood S" + [char]38 + "OP Decision Support System"
$wsAcc.Cells.Item(2,2).Value2 = $ampStr
$wsAcc.Cells.Item(2,2).Font.Size = 22
$wsAcc.Cells.Item(2,2).Font.Bold = $true
$wsAcc.Cells.Item(2,2).Font.Color = [int]0xFFFFFF
$wsAcc.Cells.Item(2,2).HorizontalAlignment = -4108
$wsAcc.Cells.Item(2,2).VerticalAlignment   = -4108
$wsAcc.Range("B2:N4").Interior.Color = [int]0x1F4975

$wsAcc.Range("B5:N5").Merge()
$wsAcc.Cells.Item(5,2).Value2 = "Systeme d'aide a la decision mensuel (Sales and Operations Planning)"
$wsAcc.Cells.Item(5,2).Font.Size = 11
$wsAcc.Cells.Item(5,2).Font.Italic = $true
$wsAcc.Cells.Item(5,2).Font.Color = [int]0x1F4975

$wsAcc.Range("B6:N6").Interior.Color = [int]0x70AD47
$wsAcc.Rows("6:6").RowHeight = 4

# Barre infos
$wsAcc.Cells.Item(8,2).Value2 = "Utilisateur :"; $wsAcc.Cells.Item(8,2).Font.Bold = $true
$wsAcc.Cells.Item(8,3).Value2 = $env:USERNAME;    $wsAcc.Cells.Item(8,3).Interior.Color = [int]0xFFFFFF; $wsAcc.Cells.Item(8,3).Borders.LineStyle = 1
$wsAcc.Cells.Item(8,5).Value2 = "Cycle :";        $wsAcc.Cells.Item(8,5).Font.Bold = $true
$wsAcc.Cells.Item(8,6).Value2 = "CYC-2023-06";    $wsAcc.Cells.Item(8,6).Interior.Color = [int]0xFFFFFF; $wsAcc.Cells.Item(8,6).Borders.LineStyle = 1
$wsAcc.Cells.Item(8,8).Value2 = "Date coupure :"; $wsAcc.Cells.Item(8,8).Font.Bold = $true
$wsAcc.Cells.Item(8,9).Value2 = "31/05/2023"
$wsAcc.Cells.Item(8,11).Value2= "Version :";      $wsAcc.Cells.Item(8,11).Font.Bold = $true
$wsAcc.Cells.Item(8,12).Value2= "1.0"

# Section workflow
$wsAcc.Range("B10:N10").Merge()
$wsAcc.Cells.Item(10,2).Value2 = "WORKFLOW S.O.P. - Cycle mensuel (utilisez les boutons ci-dessous de gauche a droite)"
$wsAcc.Cells.Item(10,2).Font.Bold = $true; $wsAcc.Cells.Item(10,2).Font.Size = 11; $wsAcc.Cells.Item(10,2).Font.Color = [int]0x1F4975
$wsAcc.Range("B10:N10").Interior.Color = [int]0xDEEBF7

# Zone statut
$wsAcc.Range("B30:N30").Merge()
$wsAcc.Cells.Item(30,2).Value2 = "STATUT :"; $wsAcc.Cells.Item(30,2).Font.Bold = $true
$wsAcc.Range("B31:N31").Merge()
$wsAcc.Cells.Item(31,2).Value2 = "En attente d'importation des donnees..."
$wsAcc.Cells.Item(31,2).Font.Italic = $true; $wsAcc.Cells.Item(31,2).Font.Color = [int]0x595959
$wsAcc.Range("B31:N31").Interior.Color = [int]0xFFFFCC

# Source donnees
$wsAcc.Cells.Item(33,2).Value2 = "Source :"; $wsAcc.Cells.Item(33,2).Font.Bold = $true
$wsAcc.Range("C33:N33").Merge()
$wsAcc.Cells.Item(33,3).Value2 = "SYNTHETIQUE - inspire de SupplyGraph (Wasi et al., 2024) | Dossier /Data"
$wsAcc.Cells.Item(33,3).Font.Italic = $true; $wsAcc.Cells.Item(33,3).Font.Color = [int]0x595959

# Named ranges pour le VBA
$wsAcc.Cells.Item(8,3).Name  = "B_USER_NAME"
$wsAcc.Cells.Item(31,2).Name = "B_STATUS"

Write-Host "   OK: ACCUEIL configure" -ForegroundColor Green

# ================================================================
Write-Host "[5/8] Injection VBA..." -ForegroundColor Yellow
$vbaProj = $wb.VBProject

$modules = @(
    "mod01_CONFIG",
    "mod02_IMPORT",
    "mod03_FORECAST",
    "mod04_PLANS",
    "mod05_SCENARIOS",
    "mod06_KPI",
    "mod07_EXPORT"
)
foreach ($modName in $modules) {
    $basPath = Join-Path $basDir "$modName.bas"
    if (-not (Test-Path $basPath)) { Write-Host "  [WARN] Manquant: $basPath" -ForegroundColor Red; continue }
    $code = Get-Content $basPath -Raw -Encoding UTF8
    $mod = $null
    foreach ($m in $vbaProj.VBComponents) { if ($m.Name -eq $modName) { $mod = $m; break } }
    if ($null -eq $mod) { $mod = $vbaProj.VBComponents.Add(1) }
    $mod.Name = $modName
    if ($mod.CodeModule.CountOfLines -gt 0) { $mod.CodeModule.DeleteLines(1, $mod.CodeModule.CountOfLines) }
    $mod.CodeModule.InsertLines(1, $code)
    Write-Host "   -> $modName OK" -ForegroundColor Gray
}

# ThisWorkbook
$twPath = Join-Path $basDir "ThisWorkbook.bas"
if (Test-Path $twPath) {
    $twCode = Get-Content $twPath -Raw -Encoding UTF8
    $tw = $vbaProj.VBComponents["ThisWorkbook"]
    if ($tw.CodeModule.CountOfLines -gt 0) { $tw.CodeModule.DeleteLines(1, $tw.CodeModule.CountOfLines) }
    $tw.CodeModule.InsertLines(1, $twCode)
    Write-Host "   -> ThisWorkbook OK" -ForegroundColor Gray
}
Write-Host "   OK: VBA injecte" -ForegroundColor Green

# ================================================================
Write-Host "[6/8] Ajout des boutons ActiveX..." -ForegroundColor Yellow
$wsAcc = $wb.Sheets["ACCUEIL"]

function AddBtn {
    param($ws, [string]$caption, [double]$left, [double]$top, [double]$width, [double]$height, [int]$bgCol)
    try {
        $ole = $ws.OLEObjects.Add("Forms.CommandButton.1",
            [System.Reflection.Missing]::Value, $false, $false,
            [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value, $left, $top, $width, $height)
        $ole.Object.Caption   = $caption
        $ole.Object.BackColor = $bgCol
        $ole.Object.ForeColor = [int]0xFFFFFF
        $ole.Object.Font.Bold = $true
        $ole.Object.Font.Size = 8
        $safe = ($caption -replace "[^A-Za-z0-9]","_") -replace "_+","_"
        $safe = $safe.Trim("_")
        $ole.Name = "btn_$safe"
        return $true
    } catch {
        Write-Host "  [WARN] Bouton '$caption': $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }
}

# Ligne 1 - J1/J2/J3
AddBtn $wsAcc "Importer les donnees"        14  220  205  42  0x2E75B6 | Out-Null
AddBtn $wsAcc "Lancer le Forecast"          230 220  180  42  0x70AD47 | Out-Null
AddBtn $wsAcc "Demand Plan"                 422 220  160  42  0x70AD47 | Out-Null
AddBtn $wsAcc "Capacity - Inventory - Gap"  594 220  230  42  0xFFC000 | Out-Null

# Ligne 2 - J4/J5/Expert
AddBtn $wsAcc "Generer les Scenarios"       14  275  200  42  0x7030A0 | Out-Null
AddBtn $wsAcc "Calculer les KPI"            226 275  170  42  0x7030A0 | Out-Null
AddBtn $wsAcc "Valider le Plan Final"       408 275  190  42  0x00B050 | Out-Null
AddBtn $wsAcc "Monte Carlo Expert"          610 275  165  42  0xC00000 | Out-Null

# Ligne 3 - Utilitaires
AddBtn $wsAcc "CYCLE COMPLET J2-J4"        14  340  190  42  0xED7D31 | Out-Null
AddBtn $wsAcc "Exporter vers Power BI"      216 340  195  42  0x595959 | Out-Null
AddBtn $wsAcc "Reinitialiser les calculs"   423 340  185  42  0xC55A11 | Out-Null
AddBtn $wsAcc "Voir le LOG"                 620 340  135  42  0x595959 | Out-Null

Write-Host "   OK: boutons ajoutes" -ForegroundColor Green

# ================================================================
Write-Host "[7/8] Liaison boutons -> macros dans Sheet ACCUEIL..." -ForegroundColor Yellow
# Injecter le code de liaison dans le module de feuille ACCUEIL
$accSheetCode = @'
Private Sub btn_Importer_les_donnees_Click()         Call ImporterDonnees:     End Sub
Private Sub btn_Lancer_le_Forecast_Click()           Call LancerForecast:      End Sub
Private Sub btn_Demand_Plan_Click()                  Call LancerDemandPlan:    End Sub
Private Sub btn_Capacity_Inventory_Gap_Click()
    Call LancerCapacityPlan
    Call LancerInventoryPlan
    Call LancerGapAnalysis
End Sub
Private Sub btn_Generer_les_Scenarios_Click()        Call GenererScenarios:    End Sub
Private Sub btn_Calculer_les_KPI_Click()             Call CalculerKPI:         End Sub
Private Sub btn_Valider_le_Plan_Final_Click()        Call ValiderPlanFinal:    End Sub
Private Sub btn_Monte_Carlo_Expert_Click()           Call LancerMonteCarlo:    End Sub
Private Sub btn_CYCLE_COMPLET_J2_J4_Click()         Call LancerCycleComplet:  End Sub
Private Sub btn_Exporter_vers_Power_BI_Click()      Call ExporterVersPowerBI:  End Sub
Private Sub btn_Reinitialiser_les_calculs_Click()   Call ReinitialiserCalculs: End Sub
Private Sub btn_Voir_le_LOG_Click()
    ThisWorkbook.Sheets("LOG").Visible = -1
    ThisWorkbook.Sheets("LOG").Activate
End Sub
'@

foreach ($vbc in $vbaProj.VBComponents) {
    $shCodeName = ""
    try { $shCodeName = $vbc.Properties.Item("Name").Value } catch {}
    if ($shCodeName -eq "ACCUEIL") {
        if ($vbc.CodeModule.CountOfLines -gt 0) { $vbc.CodeModule.DeleteLines(1, $vbc.CodeModule.CountOfLines) }
        $vbc.CodeModule.InsertLines(1, $accSheetCode)
        Write-Host "   -> Sheet ACCUEIL: code liaison OK" -ForegroundColor Gray
        break
    }
}
Write-Host "   OK: liaisons configurees" -ForegroundColor Green

# ================================================================
Write-Host "[8/8] Sauvegarde en .xlsm..." -ForegroundColor Yellow
$wb.SaveAs($xlsmPath, 52)
Write-Host "   OK: $xlsmPath" -ForegroundColor Green

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  CONSTRUCTION TERMINEE" -ForegroundColor Cyan
Write-Host "  Fichier: $xlsmPath" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

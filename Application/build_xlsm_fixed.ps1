# ================================================================
# BUILD SCRIPT â€” AtlasFood S&OP DSS
# Cree le fichier AtlasFood_SOP.xlsm complet et fonctionnel
# via Excel COM Automation (PowerShell)
# ================================================================
# PREREQUIS : Excel installe, macros autorisees
# EXECUTION  : PowerShell (en tant qu'administrateur si necessaire)
# ================================================================

param(
    [string]$WorkspaceRoot = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
)

$ErrorActionPreference = "Stop"
$xlsxPath = Join-Path $WorkspaceRoot "Application\AtlasFood_SOP.xlsm"
$dataDir  = Join-Path $WorkspaceRoot "Data"

Write-Host ""
Write-Host "========================================================"
Write-Host "  AtlasFood S&OP DSS â€” Construction du fichier .xlsm"
Write-Host "========================================================"
Write-Host ""

# --- Creer les dossiers si manquants ---
if (-not (Test-Path (Split-Path $xlsxPath))) {
    New-Item -ItemType Directory -Force -Path (Split-Path $xlsxPath) | Out-Null
}

# --- Ouvrir Excel en arriere-plan ---
Write-Host "[1/10] Demarrage d'Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

# Desactiver la protection de macro pour la construction
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow

$wb = $excel.Workbooks.Add()

# ================================================================
# CONSTANTES COULEURS
# ================================================================
$COL_HEADER_DARK  = 0x1F4975   # Bleu marine
$COL_HEADER_MED   = 0x2E75B6   # Bleu moyen
$COL_ACCENT       = 0x70AD47   # Vert
$COL_ACCENT2      = 0xED7D31   # Orange
$COL_WARN         = 0xFFC000   # Jaune
$COL_ERR          = 0xFF0000   # Rouge
$COL_WHITE        = 0xFFFFFF
$COL_LIGHTGRAY    = 0xF2F2F2
$COL_LIGHTBLUE    = 0xDEEBF7

# ================================================================
# HELPER: Ajouter ou recuperer une feuille
# ================================================================
function Get-OrAddSheet {
    param($wb, [string]$name, [string]$tabColor = "")
    $ws = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq $name) { $ws = $s; break } }
    if ($null -eq $ws) {
        $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets[$wb.Sheets.Count])
        $ws.Name = $name
    }
    if ($tabColor -ne "") {
        $ws.Tab.Color = [System.Drawing.ColorTranslator]::FromHtml($tabColor).ToArgb() -band 0xFFFFFF
    }
    return $ws
}

# ================================================================
# HELPER: Ajouter un bouton (ActiveX CommandButton) dans une feuille
# ================================================================
function Add-Button {
    param($ws, [string]$caption, [int]$left, [int]$top, [int]$width=160, [int]$height=36,
          [string]$macro="", [int]$bgColor=0x2E75B6)
    
    $btn = $ws.OLEObjects.Add("Forms.CommandButton.1", [System.Reflection.Missing]::Value,
        $false, $false, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, $left, $top, $width, $height)
    
    $btn.Object.Caption   = $caption
    $btn.Object.BackColor = $bgColor
    $btn.Object.ForeColor = 0xFFFFFF
    $btn.Object.Font.Bold = $true
    $btn.Object.Font.Size = 10
    $btn.Name = "btn_" + ($caption -replace "[^A-Za-z0-9]","_")
    
    # Le lien macro sera injecte dans le code VBA de la feuille
    return $btn
}

# ================================================================
# ETAPE 2 : CREER ET NOMMER TOUTES LES FEUILLES
# ================================================================
Write-Host "[2/10] Creation des feuilles..."

# Supprimer les feuilles par defaut sauf la premiere
while ($wb.Sheets.Count -gt 1) { $wb.Sheets[$wb.Sheets.Count].Delete() }
$wb.Sheets[1].Name = "ACCUEIL"

$shNames = @(
    @{n="ACCUEIL";        c="#1F4975"},
    @{n="PARAMETRES";     c="#70AD47"},
    @{n="PLANT";          c="#4472C4"},
    @{n="PRODUCT";        c="#4472C4"},
    @{n="CALENDAR";       c="#4472C4"},
    @{n="COST";           c="#4472C4"},
    @{n="SALES_HISTORY";  c="#ED7D31"},
    @{n="PRODUCTION_HISTORY"; c="#ED7D31"},
    @{n="DELIVERY_HISTORY";   c="#ED7D31"},
    @{n="FORECAST";       c="#FFC000"},
    @{n="DEMAND_PLAN";    c="#FFC000"},
    @{n="CAPACITY_PLAN";  c="#FFC000"},
    @{n="INVENTORY_PLAN"; c="#FFC000"},
    @{n="GAP_ANALYSIS";   c="#FF0000"},
    @{n="SCENARIOS";      c="#7030A0"},
    @{n="PLAN_FINAL";     c="#00B050"},
    @{n="KPI_RESULTS";    c="#00B050"},
    @{n="ALERTES";        c="#FF0000"},
    @{n="EXPORT_POWERBI"; c="#595959"},
    @{n="LOG";            c="#595959"}
)

foreach ($sh in $shNames) {
    $ws = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq $sh.n) { $ws = $s; break } }
    if ($null -eq $ws) {
        $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets[$wb.Sheets.Count])
        $ws.Name = $sh.n
    }
    # Couleur onglet
    $colorHex = $sh.c.TrimStart("#")
    $r = [Convert]::ToInt32($colorHex.Substring(0,2),16)
    $g = [Convert]::ToInt32($colorHex.Substring(2,2),16)
    $b = [Convert]::ToInt32($colorHex.Substring(4,2),16)
    $ws.Tab.Color = $r + ($g -shl 8) + ($b -shl 16)
}

Write-Host "   -> $($wb.Sheets.Count) feuilles creees."

# ================================================================
# ETAPE 3 : FEUILLE ACCUEIL
# ================================================================
Write-Host "[3/10] Construction de l'interface ACCUEIL..."

$wsAcc = $wb.Sheets["ACCUEIL"]
$wsAcc.Cells.ColumnWidth = 15

# Fond global
$wsAcc.Cells.Interior.Color = 0xF2F2F2

# Titre principal
$wsAcc.Rows("1:5").RowHeight = 20
$wsAcc.Range("B2:N5").Merge()
$wsAcc.Range("B2").Value = "AtlasFood S" + [char]38 + "OP Decision Support System"
$wsAcc.Range("B2").Font.Size = 26
$wsAcc.Range("B2").Font.Bold = $true
$wsAcc.Range("B2").Font.Color = 0xFFFFFF
$wsAcc.Range("B2").HorizontalAlignment = -4108  # xlCenter
$wsAcc.Range("B2").VerticalAlignment = -4108
$wsAcc.Range("B2:N5").Interior.Color = 0x1F4975

# Sous-titre
$wsAcc.Range("B6:N6").Merge()
$wsAcc.Range("B6").Value = "Systeme d'aide a la decision mensuel â€” Processus S&OP"
$wsAcc.Range("B6").Font.Size = 12
$wsAcc.Range("B6").Font.Italic = $true
$wsAcc.Range("B6").Font.Color = 0x1F4975
$wsAcc.Range("B6").HorizontalAlignment = -4108
$wsAcc.Rows("6").RowHeight = 22

# Separateur
$wsAcc.Range("B7:N7").Merge()
$wsAcc.Range("B7:N7").Interior.Color = 0x70AD47
$wsAcc.Rows("7").RowHeight = 4

# --- Zone Utilisateur ---
$wsAcc.Range("B9").Value = "Utilisateur :"
$wsAcc.Range("B9").Font.Bold = $true
$wsAcc.Range("C9").Value = Environ("USERNAME")
$wsAcc.Range("C9").Name = "B_USER_NAME"
$wsAcc.Range("C9").Interior.Color = 0xFFFFFF
$wsAcc.Range("C9").Borders.LineStyle = 1

$wsAcc.Range("E9").Value = "Cycle S&OP :"
$wsAcc.Range("E9").Font.Bold = $true
$wsAcc.Range("F9").Value = "CYC-2023-06"
$wsAcc.Range("F9").Interior.Color = 0xFFFFFF
$wsAcc.Range("F9").Borders.LineStyle = 1

$wsAcc.Range("H9").Value = "Date coupure :"
$wsAcc.Range("H9").Font.Bold = $true
$wsAcc.Range("I9").Value = "31/05/2023"
$wsAcc.Range("I9").Interior.Color = 0xFFFFFF
$wsAcc.Range("I9").Borders.LineStyle = 1

$wsAcc.Range("K9").Value = "Version outil :"
$wsAcc.Range("K9").Font.Bold = $true
$wsAcc.Range("L9").Value = "1.0"
$wsAcc.Range("L9").Interior.Color = 0xFFFFFF

# --- Section WORKFLOW ---
$wsAcc.Range("B11:N11").Merge()
$wsAcc.Range("B11").Value = "WORKFLOW S&OP â€” Cycle mensuel en 5 etapes"
$wsAcc.Range("B11").Font.Bold = $true
$wsAcc.Range("B11").Font.Size = 13
$wsAcc.Range("B11").Font.Color = 0x1F4975
$wsAcc.Rows("11").RowHeight = 22

# --- Etapes du workflow visuelles ---
$steps = @("J1 Donnees", "J2 Demande", "J3 Supply", "J4 Reunion S&OP", "J5 Decision")
$stepsCol = @("B", "D", "F", "H", "J")  # Approximation colonnes
$stepColors = @(0x2E75B6, 0x70AD47, 0xFFC000, 0xED7D31, 0x7030A0)

for ($i = 0; $i -lt 5; $i++) {
    $col = $i * 2 + 2   # colonnes B, D, F, H, J
    $wsAcc.Cells(12, $col).Value = $steps[$i]
    $wsAcc.Cells(12, $col).Font.Bold = $true
    $wsAcc.Cells(12, $col).Font.Color = 0xFFFFFF
    $wsAcc.Cells(12, $col).Interior.Color = $stepColors[$i]
    $wsAcc.Cells(12, $col).HorizontalAlignment = -4108
    $wsAcc.Rows("12").RowHeight = 24
}

# --- Section ACTIONS PRINCIPALES ---
$wsAcc.Range("B14:N14").Merge()
$wsAcc.Range("B14").Value = "ACTIONS"
$wsAcc.Range("B14").Font.Bold = $true
$wsAcc.Range("B14").Font.Size = 13
$wsAcc.Range("B14").Font.Color = 0x1F4975
$wsAcc.Range("B14:N14").Interior.Color = 0xDEEBF7
$wsAcc.Rows("14").RowHeight = 22

# --- STATUS ---
$wsAcc.Range("B30:N30").Merge()
$wsAcc.Range("B30").Value = "STATUT"
$wsAcc.Range("B30").Font.Bold = $true
$wsAcc.Range("B30").Font.Color = 0x1F4975

$wsAcc.Range("B31:N31").Merge()
$wsAcc.Range("B31").Value = "En attente d'importation des donnees..."
$wsAcc.Range("B31").Name = "B_STATUS"
$wsAcc.Range("B31").Font.Italic = $true
$wsAcc.Range("B31").Font.Color = 0x595959
$wsAcc.Range("B31:N31").Interior.Color = 0xFFFFCC

# --- Note source ---
$wsAcc.Range("B33").Value = "Source des donnees :"
$wsAcc.Range("B33").Font.Bold = $true
$wsAcc.Range("C33:N33").Merge()
$wsAcc.Range("C33").Value = "SYNTHETIQUE - inspire de SupplyGraph (ref: 2401.15299) | Dossier : /Data"
$wsAcc.Range("C33").Font.Italic = $true
$wsAcc.Range("C33").Font.Color = 0x595959

# ================================================================
# ETAPE 4 : FEUILLE PARAMETRES
# ================================================================
Write-Host "[4/10] Configuration PARAMETRES..."

$wsParam = $wb.Sheets["PARAMETRES"]
$wsParam.Cells(1,1).Value = "Cle"
$wsParam.Cells(1,2).Value = "Valeur"
$wsParam.Cells(1,3).Value = "Unite"
$wsParam.Cells(1,4).Value = "Nature"
$wsParam.Cells(1,5).Value = "Description"

# Style entete
$wsParam.Range("A1:E1").Font.Bold = $true
$wsParam.Range("A1:E1").Interior.Color = 0x1F4975
$wsParam.Range("A1:E1").Font.Color = 0xFFFFFF

$params = @(
    @("HORIZON_MOIS",        6,          "mois",   "Parametre",   "Horizon de planification (SOP)"),
    @("NIVEAU_SERVICE",      0.95,       "%",      "Parametre",   "Niveau de service cible (ex: 0.95 = 95%)"),
    @("DELAI_LIVRAISON",     5,          "jours",  "Parametre",   "Delai de reapprovisionnement moyen"),
    @("MARGE_CAPACITE",      0.10,       "%",      "Hypothese",   "Marge sur capacite observee (10pc)"),
    @("SEUIL_ALERTE_GAP",    0.85,       "%",      "Parametre",   "Seuil d'alerte utilisation capacite"),
    @("SEUIL_SURSTOCK",      2.0,        "x",      "Parametre",   "Ratio stock/demande declenchant alerte surstock"),
    @("CV_SEUIL_STABLE",     0.20,       "",       "Parametre",   "CV max pour classe STABLE"),
    @("CV_SEUIL_VOLATILE",   0.50,       "",       "Parametre",   "CV min pour classe VOLATILE"),
    @("BACKTEST_SEMAINES",   8,          "semaines","Parametre",  "Nombre de semaines pour le back-test forecast"),
    @("DATE_COUPURE",        "2023-05-31","date",  "Parametre",   "Date de coupure T0 pour le rejeu du cycle"),
    @("CYCLE_ID",            "CYC-2023-06","",     "Parametre",   "Identifiant du cycle S&OP courant"),
    @("UNITE_CHARGE",        "Tonne",    "",       "Parametre",   "Unite de charge (Tonne ou Unite)"),
    @("DEVISE",              "UM",       "",       "Hypothese",   "Unite monetaire (UM neutre = hypothese)"),
    @("STOCK_INIT_SKU01",    5000,       "unites", "Hypothese",   "Stock initial SKU-01 (ajuste pour eviter negatif)"),
    @("STOCK_INIT_SKU02",    4000,       "unites", "Hypothese",   "Stock initial SKU-02"),
    @("STOCK_INIT_SKU03",    3500,       "unites", "Hypothese",   "Stock initial SKU-03"),
    @("STOCK_INIT_SKU04",    3000,       "unites", "Hypothese",   "Stock initial SKU-04"),
    @("STOCK_INIT_SKU05",    2000,       "unites", "Hypothese",   "Stock initial SKU-05"),
    @("MC_ITERATIONS",       1000,       "",       "Parametre",   "Nombre d'iterations Monte Carlo"),
    @("MC_SIGMA_PCT",        0.15,       "%",      "Hypothese",   "Ecart-type relatif pour Monte Carlo"),
    @("SEUIL_VALIDATEUR",    "SOP Manager","",    "Parametre",   "Profil requis pour validation du plan")
)

$row = 2
foreach ($p in $params) {
    $wsParam.Cells($row, 1).Value = $p[0]
    if ([double]::TryParse($p[1], [ref]$null)) {
        $wsParam.Cells($row, 2).Value = [double]$p[1]
    } else {
        $wsParam.Cells($row, 2).Value = $p[1]
    }
    $wsParam.Cells($row, 3).Value = $p[2]
    $wsParam.Cells($row, 4).Value = $p[3]
    $wsParam.Cells($row, 5).Value = $p[4]
    
    if ($p[3] -eq "Hypothese") { $wsParam.Cells($row, 4).Font.Italic = $true }
    $row++
}

# Ajustement largeur colonnes
$wsParam.Columns("A").ColumnWidth = 22
$wsParam.Columns("B").ColumnWidth = 18
$wsParam.Columns("C").ColumnWidth = 10
$wsParam.Columns("D").ColumnWidth = 12
$wsParam.Columns("E").ColumnWidth = 55

# Convertir en tableau structure
$range = $wsParam.Range("A1:E" + ($row-1))
$lo = $wsParam.ListObjects.Add(1, $range, [System.Reflection.Missing]::Value, 1)
$lo.Name = "tbl_PARAM"
$lo.TableStyle = "TableStyleMedium2"

# ================================================================
# ETAPE 5 : FEUILLES DE DONNEES â€” EN-TETES
# ================================================================
Write-Host "[5/10] Preparation des feuilles de donnees..."

# -- PLANT --
$wsPlant = $wb.Sheets["PLANT"]
$plantHeaders = @("Plant_ID","Nom","Capacite_ref","Marge_cap","Cap_HS_max_pct","Cap_ST_max")
for ($c = 0; $c -lt $plantHeaders.Count; $c++) {
    $wsPlant.Cells(1, $c+1).Value = $plantHeaders[$c]
}
$wsPlant.Range("A1:F1").Font.Bold = $true
$wsPlant.Range("A1:F1").Interior.Color = 0x1F4975
$wsPlant.Range("A1:F1").Font.Color = 0xFFFFFF
$wsPlant.Range("A2").Value = "USN-01"; $wsPlant.Range("B2").Value = "Usine Nord"
$wsPlant.Range("C2").Value = 5000; $wsPlant.Range("D2").Value = 0.10
$wsPlant.Range("E2").Value = 0.20; $wsPlant.Range("F2").Value = 1000
$wsPlant.Range("A3").Value = "USN-02"; $wsPlant.Range("B3").Value = "Usine Sud"
$wsPlant.Range("C3").Value = 4000; $wsPlant.Range("D3").Value = 0.10
$wsPlant.Range("E3").Value = 0.20; $wsPlant.Range("F3").Value = 800
$wsPlant.Range("A4").Value = "USN-03"; $wsPlant.Range("B4").Value = "Usine Est"
$wsPlant.Range("C4").Value = 2500; $wsPlant.Range("D4").Value = 0.10
$wsPlant.Range("E4").Value = 0.15; $wsPlant.Range("F4").Value = 500
$lo = $wsPlant.ListObjects.Add(1, $wsPlant.Range("A1:F4"), [System.Reflection.Missing]::Value, 1)
$lo.Name = "tbl_PLANT"; $lo.TableStyle = "TableStyleMedium2"

# -- PRODUCT --
$wsProd = $wb.Sheets["PRODUCT"]
$prodHeaders = @("SKU_ID","Nom","Famille","Plant_ID","Poids_unit","DLC_jours","Stock_secu_methode")
for ($c = 0; $c -lt $prodHeaders.Count; $c++) { $wsProd.Cells(1,$c+1).Value = $prodHeaders[$c] }
$wsProd.Range("A1:G1").Font.Bold = $true
$wsProd.Range("A1:G1").Interior.Color = 0x1F4975
$wsProd.Range("A1:G1").Font.Color = 0xFFFFFF
$prodData = @(
    @("SKU-01","Biscuits Nature 250g","Biscuits","USN-01",0.25,180,"SS_NORMAL"),
    @("SKU-02","Biscuits Chocolat 250g","Biscuits","USN-01",0.25,150,"SS_NORMAL"),
    @("SKU-03","Crackers Sale 200g","Crackers","USN-02",0.20,120,"SS_NORMAL"),
    @("SKU-04","Crackers Cereales 200g","Crackers","USN-02",0.20,120,"SS_NORMAL"),
    @("SKU-05","Gaufrettes Vanille 150g","Gaufrettes","USN-03",0.15,90,"SS_SEASONAL")
)
for ($r = 0; $r -lt $prodData.Count; $r++) {
    for ($c = 0; $c -lt $prodData[$r].Count; $c++) {
        $wsProd.Cells($r+2, $c+1).Value = $prodData[$r][$c]
    }
}
$lo = $wsProd.ListObjects.Add(1, $wsProd.Range("A1:G" + ($prodData.Count+1)), [System.Reflection.Missing]::Value, 1)
$lo.Name = "tbl_PRODUCT"; $lo.TableStyle = "TableStyleMedium2"

# -- CALENDAR --
$wsCal = $wb.Sheets["CALENDAR"]
$calHeaders = @("Mois","Jours_ouvres","Jours_feries","Jours_Ramadan")
for ($c = 0; $c -lt $calHeaders.Count; $c++) { $wsCal.Cells(1,$c+1).Value = $calHeaders[$c] }
$wsCal.Range("A1:D1").Font.Bold = $true
$wsCal.Range("A1:D1").Interior.Color = 0x1F4975
$wsCal.Range("A1:D1").Font.Color = 0xFFFFFF
$calData = @(
    @("2023-01-01",22,0,0),@("2023-02-01",20,0,0),@("2023-03-01",23,1,22),
    @("2023-04-01",18,2,8),@("2023-05-01",22,1,0),@("2023-06-01",21,1,0),
    @("2023-07-01",21,0,0),@("2023-08-01",7,0,0)
)
for ($r = 0; $r -lt $calData.Count; $r++) {
    $wsCal.Cells($r+2,1).Value = $calData[$r][0]
    $wsCal.Cells($r+2,2).Value = $calData[$r][1]
    $wsCal.Cells($r+2,3).Value = $calData[$r][2]
    $wsCal.Cells($r+2,4).Value = $calData[$r][3]
}
$lo = $wsCal.ListObjects.Add(1, $wsCal.Range("A1:D" + ($calData.Count+1)), [System.Reflection.Missing]::Value, 1)
$lo.Name = "tbl_CALENDAR"; $lo.TableStyle = "TableStyleMedium2"

# -- COST --
$wsCost = $wb.Sheets["COST"]
$costHeaders = @("Type","Plant_ID","Montant_unit","Unite_monet")
for ($c = 0; $c -lt $costHeaders.Count; $c++) { $wsCost.Cells(1,$c+1).Value = $costHeaders[$c] }
$wsCost.Range("A1:D1").Font.Bold = $true
$wsCost.Range("A1:D1").Interior.Color = 0x1F4975
$wsCost.Range("A1:D1").Font.Color = 0xFFFFFF
$costData = @(
    @("HS","USN-01",350,"UM"),@("HS","USN-02",340,"UM"),@("HS","USN-03",320,"UM"),
    @("ST","USN-01",500,"UM"),@("ST","USN-02",490,"UM"),@("ST","USN-03",480,"UM"),
    @("Stockage","USN-01",12,"UM"),@("Stockage","USN-02",12,"UM"),@("Stockage","USN-03",11,"UM"),
    @("Penurie","USN-01",800,"UM"),@("Penurie","USN-02",800,"UM"),@("Penurie","USN-03",750,"UM"),
    @("Peremption","USN-01",200,"UM"),@("Peremption","USN-02",200,"UM"),@("Peremption","USN-03",190,"UM"),
    @("Prod","USN-01",150,"UM"),@("Prod","USN-02",145,"UM"),@("Prod","USN-03",140,"UM")
)
for ($r = 0; $r -lt $costData.Count; $r++) {
    for ($c = 0; $c -lt 4; $c++) { $wsCost.Cells($r+2,$c+1).Value = $costData[$r][$c] }
}
$lo = $wsCost.ListObjects.Add(1, $wsCost.Range("A1:D" + ($costData.Count+1)), [System.Reflection.Missing]::Value, 1)
$lo.Name = "tbl_COST"; $lo.TableStyle = "TableStyleMedium2"

# -- En-tetes feuilles donnees historiques (vides, remplies par VBA import) --
$histSheets = @{
    "SALES_HISTORY"       = @("SKU_ID","Date","Qte_commandee","Unite");
    "PRODUCTION_HISTORY"  = @("SKU_ID","Plant_ID","Date","Qte_produite");
    "DELIVERY_HISTORY"    = @("SKU_ID","Date","Qte_livree");
}
$histTableNames = @{
    "SALES_HISTORY"       = "tbl_SALES";
    "PRODUCTION_HISTORY"  = "tbl_PROD_HIST";
    "DELIVERY_HISTORY"    = "tbl_DELIV";
}
foreach ($shName in $histSheets.Keys) {
    $wsH = $wb.Sheets[$shName]
    $hdrs = $histSheets[$shName]
    for ($c = 0; $c -lt $hdrs.Count; $c++) { $wsH.Cells(1,$c+1).Value = $hdrs[$c] }
    $lastCol = [char](64 + $hdrs.Count)
    $wsH.Range("A1:${lastCol}1").Font.Bold = $true
    $wsH.Range("A1:${lastCol}1").Interior.Color = 0xED7D31
    $wsH.Range("A1:${lastCol}1").Font.Color = 0xFFFFFF
    # Ajouter une ligne vide pour pouvoir creer le tableau
    $wsH.Cells(2,1).Value = "(vide - utilisez le bouton Importer)"
    $wsH.Cells(2,1).Font.Italic = $true
    $wsH.Cells(2,1).Font.Color = 0x595959
}

# -- En-tetes feuilles calculees --
# FORECAST
$wsFc = $wb.Sheets["FORECAST"]
$fcHdrs = @("SKU_ID","Periode","Qte_prevue","Methode","WAPE_pct","Biais_pct","Statut","Cycle_ID","Classe")
for ($c = 0; $c -lt $fcHdrs.Count; $c++) { $wsFc.Cells(1,$c+1).Value = $fcHdrs[$c] }
$wsFc.Range("A1:I1").Font.Bold = $true
$wsFc.Range("A1:I1").Interior.Color = 0x1F4975
$wsFc.Range("A1:I1").Font.Color = 0xFFFFFF

# DEMAND_PLAN
$wsDp = $wb.Sheets["DEMAND_PLAN"]
$dpHdrs = @("SKU_ID","Mois","Qte_stat","Ajustement","Qte_finale","Cycle_ID","Commentaire")
for ($c = 0; $c -lt $dpHdrs.Count; $c++) { $wsDp.Cells(1,$c+1).Value = $dpHdrs[$c] }
$wsDp.Range("A1:G1").Font.Bold = $true
$wsDp.Range("A1:G1").Interior.Color = 0x1F4975
$wsDp.Range("A1:G1").Font.Color = 0xFFFFFF

# CAPACITY_PLAN
$wsCap = $wb.Sheets["CAPACITY_PLAN"]
$capHdrs = @("Plant_ID","Mois","Capacite_normale","Cap_HS_max","Cap_ST_max","Charge","Utilisation_pct","Statut","Cycle_ID")
for ($c = 0; $c -lt $capHdrs.Count; $c++) { $wsCap.Cells(1,$c+1).Value = $capHdrs[$c] }
$wsCap.Range("A1:I1").Font.Bold = $true
$wsCap.Range("A1:I1").Interior.Color = 0x1F4975
$wsCap.Range("A1:I1").Font.Color = 0xFFFFFF

# INVENTORY_PLAN
$wsInv = $wb.Sheets["INVENTORY_PLAN"]
$invHdrs = @("SKU_ID","Mois","Stock_debut","Stock_secu","Production","Livraisons","Stock_fin","DLC_restante_est","Statut","Cycle_ID")
for ($c = 0; $c -lt $invHdrs.Count; $c++) { $wsInv.Cells(1,$c+1).Value = $invHdrs[$c] }
$wsInv.Range("A1:J1").Font.Bold = $true
$wsInv.Range("A1:J1").Interior.Color = 0x1F4975
$wsInv.Range("A1:J1").Font.Color = 0xFFFFFF

# GAP_ANALYSIS
$wsGap = $wb.Sheets["GAP_ANALYSIS"]
$gapHdrs = @("Plant_ID","Mois","Besoin_net","Cap_disponible","Gap","Gap_pct","Statut","Cycle_ID")
for ($c = 0; $c -lt $gapHdrs.Count; $c++) { $wsGap.Cells(1,$c+1).Value = $gapHdrs[$c] }
$wsGap.Range("A1:H1").Font.Bold = $true
$wsGap.Range("A1:H1").Interior.Color = 0x1F4975
$wsGap.Range("A1:H1").Font.Color = 0xFFFFFF

# SCENARIOS
$wsSc = $wb.Sheets["SCENARIOS"]
$scHdrs = @("Scenario_ID","Nom","Hypotheses","Cout_total","Taux_service_pct","Stock_final_tot","Risque_peremption","Util_max_pct","Cycle_ID","Statut")
for ($c = 0; $c -lt $scHdrs.Count; $c++) { $wsSc.Cells(1,$c+1).Value = $scHdrs[$c] }
$wsSc.Range("A1:J1").Font.Bold = $true
$wsSc.Range("A1:J1").Interior.Color = 0x1F4975
$wsSc.Range("A1:J1").Font.Color = 0xFFFFFF

# PLAN_FINAL
$wsPf = $wb.Sheets["PLAN_FINAL"]
$pfHdrs = @("Plan_ID","Version","SKU_ID","Plant_ID","Mois","Prod_normale","Prod_HS","Prod_ST","Stock_fin","Demande","Non_servi","Scenario_retenu","Statut","Valideur","Horodatage")
for ($c = 0; $c -lt $pfHdrs.Count; $c++) { $wsPf.Cells(1,$c+1).Value = $pfHdrs[$c] }
$wsPf.Range("A1:O1").Font.Bold = $true
$wsPf.Range("A1:O1").Interior.Color = 0x1F4975
$wsPf.Range("A1:O1").Font.Color = 0xFFFFFF

# KPI_RESULTS
$wsKpi = $wb.Sheets["KPI_RESULTS"]
$kpiHdrs = @("KPI_ID","Indicateur","Perimetre","Periode","Valeur","Unite","Seuil_vert","Seuil_rouge","Statut_seuil","Cycle_ID")
for ($c = 0; $c -lt $kpiHdrs.Count; $c++) { $wsKpi.Cells(1,$c+1).Value = $kpiHdrs[$c] }
$wsKpi.Range("A1:J1").Font.Bold = $true
$wsKpi.Range("A1:J1").Interior.Color = 0x1F4975
$wsKpi.Range("A1:J1").Font.Color = 0xFFFFFF

# ALERTES
$wsAlt = $wb.Sheets["ALERTES"]
$altHdrs = @("Alert_ID","Type","Criticite","Perimetre","Message","Cycle_ID","Horodatage")
for ($c = 0; $c -lt $altHdrs.Count; $c++) { $wsAlt.Cells(1,$c+1).Value = $altHdrs[$c] }
$wsAlt.Range("A1:G1").Font.Bold = $true
$wsAlt.Range("A1:G1").Interior.Color = 0xFF0000
$wsAlt.Range("A1:G1").Font.Color = 0xFFFFFF

# EXPORT_POWERBI
$wsExp = $wb.Sheets["EXPORT_POWERBI"]
$wsExp.Range("A1").Value = "Table"
$wsExp.Range("B1").Value = "Derniere_export"
$wsExp.Range("C1").Value = "Nb_lignes"
$wsExp.Range("D1").Value = "Statut"
$wsExp.Range("A1:D1").Font.Bold = $true
$wsExp.Range("A1:D1").Interior.Color = 0x595959
$wsExp.Range("A1:D1").Font.Color = 0xFFFFFF

# LOG
$wsLog = $wb.Sheets["LOG"]
$logHdrs = @("Horodatage","Utilisateur","Module","Niveau","Message")
for ($c = 0; $c -lt $logHdrs.Count; $c++) { $wsLog.Cells(1,$c+1).Value = $logHdrs[$c] }
$wsLog.Range("A1:E1").Font.Bold = $true
$wsLog.Range("A1:E1").Interior.Color = 0x595959
$wsLog.Range("A1:E1").Font.Color = 0xFFFFFF

Write-Host "   -> Toutes les feuilles configurees."

# ================================================================
# ETAPE 6 : INJECTION VBA
# ================================================================
Write-Host "[6/10] Injection du code VBA..."

$vbaProject = $wb.VBProject

# Fonction helper pour ajouter/remplacer un module VBA
function Set-VBAModule {
    param($vbaProject, [string]$name, [string]$code, [int]$type = 1)
    # type: 1=module standard, 2=class module
    $mod = $null
    foreach ($m in $vbaProject.VBComponents) {
        if ($m.Name -eq $name) { $mod = $m; break }
    }
    if ($null -eq $mod) {
        $mod = $vbaProject.VBComponents.Add($type)
        $mod.Name = $name
    }
    $mod.CodeModule.DeleteLines(1, $mod.CodeModule.CountOfLines)
    $mod.CodeModule.InsertLines(1, $code)
}

# ============================================================
# MODULE 01 : CONFIG
# ============================================================
$vba_CONFIG = @'
Option Explicit
'==============================================================
' MODULE CONFIG - Parametres centraux AtlasFood S&OP DSS v1.0
'==============================================================

Public Const TOOL_VERSION   As String = "1.0"
Public Const TOOL_NAME      As String = "AtlasFood S&OP DSS"

' Noms des feuilles
Public Const SH_ACCUEIL     As String = "ACCUEIL"
Public Const SH_PARAM       As String = "PARAMETRES"
Public Const SH_PLANT       As String = "PLANT"
Public Const SH_PRODUCT     As String = "PRODUCT"
Public Const SH_CALENDAR    As String = "CALENDAR"
Public Const SH_COST        As String = "COST"
Public Const SH_SALES       As String = "SALES_HISTORY"
Public Const SH_PROD_HIST   As String = "PRODUCTION_HISTORY"
Public Const SH_DELIV       As String = "DELIVERY_HISTORY"
Public Const SH_FORECAST    As String = "FORECAST"
Public Const SH_DEMAND_PLAN As String = "DEMAND_PLAN"
Public Const SH_CAPACITY    As String = "CAPACITY_PLAN"
Public Const SH_INVENTORY   As String = "INVENTORY_PLAN"
Public Const SH_GAP         As String = "GAP_ANALYSIS"
Public Const SH_SCENARIOS   As String = "SCENARIOS"
Public Const SH_PLAN_FINAL  As String = "PLAN_FINAL"
Public Const SH_KPI         As String = "KPI_RESULTS"
Public Const SH_ALERTS      As String = "ALERTES"
Public Const SH_EXPORT      As String = "EXPORT_POWERBI"
Public Const SH_LOG         As String = "LOG"

'--------------------------------------------------------------
' Lire un parametre depuis la feuille PARAMETRES
'--------------------------------------------------------------
Public Function GetParam(key As String) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cell As Range
    On Error GoTo Defaut
    Set ws = ThisWorkbook.Sheets(SH_PARAM)
    Set lo = ws.ListObjects("tbl_PARAM")
    For Each cell In lo.ListColumns("Cle").DataBodyRange
        If cell.Value = key Then
            GetParam = cell.Offset(0, 1).Value
            Exit Function
        End If
    Next cell
Defaut:
    Select Case key
        Case "HORIZON_MOIS":      GetParam = 6
        Case "NIVEAU_SERVICE":    GetParam = 0.95
        Case "DELAI_LIVRAISON":   GetParam = 5
        Case "MARGE_CAPACITE":    GetParam = 0.1
        Case "SEUIL_ALERTE_GAP":  GetParam = 0.85
        Case "SEUIL_SURSTOCK":    GetParam = 2
        Case "CV_SEUIL_STABLE":   GetParam = 0.2
        Case "CV_SEUIL_VOLATILE": GetParam = 0.5
        Case "BACKTEST_SEMAINES": GetParam = 8
        Case "DATE_COUPURE":      GetParam = "2023-05-31"
        Case "CYCLE_ID":          GetParam = "CYC-2023-06"
        Case "MC_ITERATIONS":     GetParam = 1000
        Case "MC_SIGMA_PCT":      GetParam = 0.15
        Case Else:                GetParam = 0
    End Select
End Function

'--------------------------------------------------------------
' Journalisation
'--------------------------------------------------------------
Public Sub LogMsg(module_ As String, niveau As String, message As String)
    Dim ws As Worksheet
    Dim r As Long
    On Error Exit Sub
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now()
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value = GetCurrentUser()
    ws.Cells(r, 3).Value = module_
    ws.Cells(r, 4).Value = niveau
    ws.Cells(r, 5).Value = message
    Select Case UCase(niveau)
        Case "ERROR":   ws.Cells(r, 4).Interior.Color = RGB(255, 100, 100)
        Case "WARNING": ws.Cells(r, 4).Interior.Color = RGB(255, 200, 80)
        Case Else:      ws.Cells(r, 4).Interior.Color = RGB(200, 230, 255)
    End Select
End Sub

Public Function GetCurrentUser() As String
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Sheets(SH_ACCUEIL).Range("B_USER_NAME").Value
    If v <> "" And Not IsEmpty(v) Then GetCurrentUser = CStr(v): Exit Function
    GetCurrentUser = Environ("USERNAME")
End Function

Public Function GetCycleID() As String
    GetCycleID = CStr(GetParam("CYCLE_ID"))
    If GetCycleID = "" Then GetCycleID = "CYC-" & Format(Now(), "YYYY-MM")
End Function

Public Function GetHorizon() As Integer
    Dim h As Variant: h = GetParam("HORIZON_MOIS")
    If IsNumeric(h) Then GetHorizon = CInt(h) Else GetHorizon = 6
End Function

Public Function GetDateCoupure() As Date
    On Error Resume Next
    GetDateCoupure = CDate(GetParam("DATE_COUPURE"))
    If Err.Number <> 0 Then GetDateCoupure = DateSerial(2023, 5, 31)
End Function

Public Sub SetStatus(msg As String)
    On Error Resume Next
    ThisWorkbook.Sheets(SH_ACCUEIL).Range("B_STATUS").Value = msg
End Sub

Public Sub ShowProgress(step_ As String, pct As Integer)
    Application.StatusBar = "[" & TOOL_NAME & "] " & step_ & " - " & pct & "%"
    DoEvents
End Sub

Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' Effacer le contenu d'une feuille (sauf ligne 1 = en-tetes)
Public Sub ClearSheet(shName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    If ws Is Nothing Then Exit Sub
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range("A2:" & ws.Cells(ws.Rows.Count, ws.Columns.Count - 1).Address).ClearContents
        ws.Range("A2:" & ws.Cells(ws.Rows.Count, ws.Columns.Count - 1).Address).Interior.ColorIndex = xlNone
    End If
End Sub
'@

Set-VBAModule -vbaProject $vbaProject -name "mod01_CONFIG" -code $vba_CONFIG

# ============================================================
# MODULE 02 : IMPORT
# ============================================================
$vba_IMPORT = @'
Option Explicit
'==============================================================
' MODULE IMPORT - Importation CSV + validation donnees
' F1 / EF-01 â€” AtlasFood S&OP DSS
'==============================================================

Public Sub ImporterDonnees()
    Dim dataDir As String
    Dim fd As FileDialog
    Dim errCount As Integer
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "IMPORT", "INFO", "=== DEBUT IMPORTATION ==="
    ShowProgress "Choix du dossier...", 0
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Selectionner le dossier Data (contenant les fichiers CSV)"
    fd.InitialFileName = ThisWorkbook.Path
    
    If fd.Show = -1 Then
        dataDir = fd.SelectedItems(1)
    Else
        MsgBox "Importation annulee.", vbInformation, TOOL_NAME
        GoTo CleanExit
    End If
    
    errCount = 0
    
    ShowProgress "Import SALES_HISTORY...", 20
    If Not ImporterCSV(dataDir & "\SALES_HISTORY.csv", SH_SALES, "tbl_SALES") Then errCount = errCount + 1
    
    ShowProgress "Import PRODUCTION_HISTORY...", 45
    If Not ImporterCSV(dataDir & "\PRODUCTION_HISTORY.csv", SH_PROD_HIST, "tbl_PROD_HIST") Then errCount = errCount + 1
    
    ShowProgress "Import DELIVERY_HISTORY...", 70
    If Not ImporterCSV(dataDir & "\DELIVERY_HISTORY.csv", SH_DELIV, "tbl_DELIV") Then errCount = errCount + 1
    
    ShowProgress "Validation...", 90
    Dim warns As Integer
    warns = ValiderDonnees()
    
    ClearProgress
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim msg As String
    If errCount = 0 Then
        msg = "Import reussi !" & vbCrLf & _
              warns & " avertissement(s)." & vbCrLf & vbCrLf & _
              "Vous pouvez lancer le Forecast (J2 Demande)."
        MsgBox msg, vbInformation, TOOL_NAME
        SetStatus "Donnees importees â€” " & Now() & " â€” " & warns & " avertissement(s)"
        LogMsg "IMPORT", "INFO", "Import OK - " & warns & " avertissement(s)"
    Else
        MsgBox errCount & " erreur(s) lors de l'import." & vbCrLf & "Consultez l'onglet LOG.", vbCritical, TOOL_NAME
        SetStatus "ERREUR d'importation â€” consultez LOG"
        LogMsg "IMPORT", "ERROR", errCount & " erreur(s)"
    End If
    Exit Sub
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    LogMsg "IMPORT", "ERROR", "Erreur fatale: " & Err.Description
    MsgBox "Erreur: " & Err.Description, vbCritical, TOOL_NAME
End Sub

' Importer un CSV (UTF-8 via ADODB.Stream) dans une feuille
Private Function ImporterCSV(filePath As String, shName As String, tblName As String) As Boolean
    ImporterCSV = False
    
    If Dir(filePath) = "" Then
        LogMsg "IMPORT", "ERROR", "Fichier manquant: " & filePath
        Exit Function
    End If
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then LogMsg "IMPORT", "ERROR", "Feuille inconnue: " & shName: Exit Function
    
    ' Lire via ADODB.Stream (UTF-8)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    Dim allText As String
    allText = stm.ReadText()
    stm.Close
    
    Dim allLines() As String
    allLines = Split(allText, vbLf)
    
    ' Effacer l'ancienne data (garder entete ligne 1)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR > 1 Then ws.Rows("2:" & lastR).Delete
    
    ' Lire en-tetes du CSV (ligne 0)
    Dim hdrs() As String
    hdrs = Split(Replace(allLines(0), vbCr, ""), ";")
    Dim j As Integer
    For j = 0 To UBound(hdrs): hdrs(j) = Replace(hdrs(j), Chr(34), ""): hdrs(j) = Trim(hdrs(j)): Next j
    
    ' Ecrire les donnees
    Dim row As Long: row = 2
    Dim k As Long
    For k = 1 To UBound(allLines)
        Dim lineStr As String
        lineStr = Replace(allLines(k), vbCr, "")
        If Trim(lineStr) = "" Then GoTo NextL
        
        Dim vals() As String
        vals = Split(lineStr, ";")
        
        Dim col As Integer
        For col = 0 To UBound(vals)
            If col > 30 Then Exit For
            Dim v As String
            v = Replace(vals(col), Chr(34), "")
            v = Trim(v)
            
            If IsNumeric(v) Then
                ws.Cells(row, col + 1).Value = CDbl(v)
            ElseIf Len(v) = 10 And Mid(v, 5, 1) = "-" And Mid(v, 8, 1) = "-" Then
                ws.Cells(row, col + 1).Value = CDate(v)
                ws.Cells(row, col + 1).NumberFormat = "yyyy-mm-dd"
            Else
                ws.Cells(row, col + 1).Value = v
            End If
        Next col
        row = row + 1
NextL:
    Next k
    
    Dim nbRows As Long: nbRows = row - 2
    
    ' Creer / recreer le tableau structure
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    
    Dim lastCol As Integer: lastCol = UBound(hdrs) + 1
    If lo Is Nothing And nbRows > 0 Then
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(row - 1, lastCol)), , xlYes)
        lo.Name = tblName
        lo.TableStyle = "TableStyleMedium9"
    End If
    
    LogMsg "IMPORT", "INFO", "OK: " & shName & " (" & nbRows & " lignes)"
    ImporterCSV = True
End Function

' Validation croisee (BR-01 a BR-06)
Private Function ValiderDonnees() As Integer
    Dim warns As Integer: warns = 0
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cell As Range
    Dim avg As Double, stdv As Double
    
    ' BR-01 + BR-02: Valeurs manquantes et negatives dans SALES
    Set ws = ThisWorkbook.Sheets(SH_SALES)
    On Error Resume Next
    Set lo = ws.ListObjects("tbl_SALES")
    On Error GoTo 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Dim qteCol As Range
            Set qteCol = lo.ListColumns("Qte_commandee").DataBodyRange
            avg = Application.Average(qteCol)
            stdv = Application.StDev(qteCol)
            For Each cell In qteCol
                If IsEmpty(cell.Value) Or cell.Value = "" Then
                    LogMsg "VALID", "WARNING", "BR-01: Manquant SALES ligne " & cell.Row
                    cell.Interior.Color = RGB(255, 255, 150)
                    warns = warns + 1
                ElseIf IsNumeric(cell.Value) Then
                    If cell.Value < 0 Then
                        LogMsg "VALID", "WARNING", "BR-02: Negatif SALES ligne " & cell.Row
                        cell.Interior.Color = RGB(255, 150, 150)
                        warns = warns + 1
                    ElseIf cell.Value > avg + 3 * stdv And stdv > 0 Then
                        LogMsg "VALID", "WARNING", "BR-06: Outlier SALES ligne " & cell.Row & " val=" & cell.Value
                        warns = warns + 1
                    End If
                End If
            Next cell
        End If
    End If
    
    ' BR-04: Aout 2023 partiel
    LogMsg "VALID", "INFO", "BR-04: Mois partiel Aout 2023 (9 jours) â€” exclu des calculs mensuels"
    
    LogMsg "VALID", "INFO", "Validation OK â€” " & warns & " avertissement(s)"
    ValiderDonnees = warns
End Function
'@

Set-VBAModule -vbaProject $vbaProject -name "mod02_IMPORT" -code $vba_IMPORT

# ============================================================
# MODULE 03 : FORECAST
# ============================================================
$vba_FORECAST = @'
Option Explicit
'==============================================================
' MODULE FORECAST - Prevision multi-methodes + back-test
' F2 (classification) + F3 (forecast) / EF-02 + EF-03
' Methodes: Moyenne Mobile, Lissage Exponentiel, Holt, Holt-Winters simplifie
'==============================================================

Public Sub LancerForecast()
    Dim cycleID As String
    Dim dateCoup As Date
    Dim horizon  As Integer
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "FORECAST", "INFO", "=== DEBUT FORECAST ==="
    ShowProgress "Preparation...", 5
    SetStatus "Forecast en cours..."
    
    cycleID  = GetCycleID()
    dateCoup = GetDateCoupure()
    horizon  = GetHorizon()
    
    ' Verifier donnees
    Dim wsSales As Worksheet
    Set wsSales = ThisWorkbook.Sheets(SH_SALES)
    If wsSales.Cells(wsSales.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Aucune donnee dans SALES_HISTORY." & vbCrLf & "Importez d'abord les donnees (J1).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ' Effacer forecast precedent
    Dim wsFc As Worksheet
    Set wsFc = ThisWorkbook.Sheets(SH_FORECAST)
    ClearSheet SH_FORECAST
    
    ' Recuperer liste SKU
    Dim skus() As String
    skus = GetSKUList()
    If UBound(skus) < 0 Then
        MsgBox "Aucun produit dans la table PRODUCT.", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    Dim rowOut As Long: rowOut = 2
    Dim i As Integer
    
    For i = 0 To UBound(skus)
        Dim sku As String: sku = skus(i)
        ShowProgress "Forecast: " & sku, 10 + Int(85 * i / (UBound(skus) + 1))
        
        ' Historique mensuel jusqu'a la date de coupure
        Dim histD() As Double
        Dim histP() As Date
        histD = HistMensuel(sku, dateCoup, histP)
        
        If UBound(histD) < 2 Then
            LogMsg "FORECAST", "WARNING", "Histo insuffisant: " & sku
            GoTo SuivantSKU
        End If
        
        ' Classification F2
        Dim classe As String: classe = ClasserProduit(histD)
        
        ' 4 methodes de forecast
        Dim btLen As Integer: btLen = 2  ' 2 mois backtest
        Dim wape0 As Double, wape1 As Double, wape2 As Double, wape3 As Double
        Dim biais0 As Double, biais1 As Double, biais2 As Double, biais3 As Double
        Dim fc0(5) As Double, fc1(5) As Double, fc2(5) As Double, fc3(5) As Double
        
        Call MoyenneMobile3(histD, horizon, btLen, fc0, wape0, biais0)
        Call LissageExpo(histD, horizon, btLen, fc1, wape1, biais1)
        Call Holt(histD, horizon, btLen, fc2, wape2, biais2)
        Call HoltWinters(histD, horizon, btLen, fc3, wape3, biais3)
        
        ' Meilleure methode = WAPE minimum
        Dim bestWAPE As Double: bestWAPE = wape0
        Dim bestIdx  As Integer: bestIdx = 0
        If wape1 < bestWAPE Then bestWAPE = wape1: bestIdx = 1
        If wape2 < bestWAPE Then bestWAPE = wape2: bestIdx = 2
        If wape3 < bestWAPE Then bestWAPE = wape3: bestIdx = 3
        
        ' Ecrire les 4 methodes (methode retenue mise en evidence)
        Dim mStart As Date
        mStart = DateSerial(Year(DateAdd("m", 1, dateCoup)), Month(DateAdd("m", 1, dateCoup)), 1)
        
        Dim m As Integer, p As Integer
        Dim wapes(3) As Double: wapes(0)=wape0: wapes(1)=wape1: wapes(2)=wape2: wapes(3)=wape3
        Dim biaiss(3) As Double: biaiss(0)=biais0: biaiss(1)=biais1: biaiss(2)=biais2: biaiss(3)=biais3
        Dim methodes(3) As String
        methodes(0)="Moyenne Mobile 3": methodes(1)="Lissage Exponentiel"
        methodes(2)="Holt Tendance":    methodes(3)="Holt-Winters Saison"
        
        For m = 0 To 3
            For p = 0 To horizon - 1
                Dim mDate As Date: mDate = DateAdd("m", p, mStart)
                Dim statut As String
                If m = bestIdx Then statut = "RETENUE" Else statut = "Alternative"
                
                Dim fcVal As Double
                Select Case m
                    Case 0: fcVal = fc0(p)
                    Case 1: fcVal = fc1(p)
                    Case 2: fcVal = fc2(p)
                    Case 3: fcVal = fc3(p)
                End Select
                
                wsFc.Cells(rowOut, 1).Value = sku
                wsFc.Cells(rowOut, 2).Value = mDate
                wsFc.Cells(rowOut, 2).NumberFormat = "mmm-yy"
                wsFc.Cells(rowOut, 3).Value = Round(Application.Max(0, fcVal), 0)
                wsFc.Cells(rowOut, 4).Value = methodes(m)
                wsFc.Cells(rowOut, 5).Value = Round(wapes(m) * 100, 1)
                wsFc.Cells(rowOut, 6).Value = Round(biaiss(m) * 100, 1)
                wsFc.Cells(rowOut, 7).Value = statut
                wsFc.Cells(rowOut, 8).Value = cycleID
                wsFc.Cells(rowOut, 9).Value = classe
                
                If statut = "RETENUE" Then
                    wsFc.Rows(rowOut).Interior.Color = RGB(200, 240, 200)
                Else
                    wsFc.Rows(rowOut).Interior.ColorIndex = xlNone
                End If
                rowOut = rowOut + 1
            Next p
        Next m
        LogMsg "FORECAST", "INFO", sku & " => " & classe & " | Methode: " & methodes(bestIdx) & " | WAPE=" & Round(bestWAPE*100,1) & "%"
SuivantSKU:
    Next i
    
    ' Creer tableau structure
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next
        Set lo = wsFc.ListObjects("tbl_FORECAST")
        On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsFc.ListObjects.Add(xlSrcRange, wsFc.Range("A1:I" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_FORECAST"
            lo.TableStyle = "TableStyleMedium3"
        End If
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    SetStatus "Forecast termine â€” " & Now()
    LogMsg "FORECAST", "INFO", "Forecast termine. " & (UBound(skus)+1) & " SKU."
    
    MsgBox "Forecast calcule pour " & (UBound(skus)+1) & " produits." & vbCrLf & _
           "Consultez l'onglet FORECAST." & vbCrLf & vbCrLf & _
           "Etape suivante : lancer le Demand Plan (J2).", vbInformation, TOOL_NAME
    ThisWorkbook.Sheets(SH_FORECAST).Activate
    Exit Sub
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    LogMsg "FORECAST", "ERROR", Err.Description
    MsgBox "Erreur Forecast: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'--- Historique mensuel agreage jusqu'a dateCoupure --
Private Function HistMensuel(sku As String, cutDate As Date, ByRef periods() As Date) As Double()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_SALES)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tbl_SALES"): On Error GoTo 0
    
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        ReDim HistMensuel(-1 To -1): Exit Function
    End If
    
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If CStr(lo.DataBodyRange(r, 1).Value) <> sku Then GoTo NL
        Dim dv As Variant: dv = lo.DataBodyRange(r, 2).Value
        If Not IsDate(dv) Then GoTo NL
        If CDate(dv) > cutDate Then GoTo NL
        Dim mKey As String: mKey = Format(CDate(dv), "YYYY-MM")
        Dim qty As Double
        If IsNumeric(lo.DataBodyRange(r, 3).Value) Then qty = CDbl(lo.DataBodyRange(r, 3).Value)
        If dict.Exists(mKey) Then dict(mKey) = dict(mKey) + qty Else dict.Add mKey, qty
NL:
    Next r
    
    Dim n As Integer: n = dict.Count
    If n = 0 Then ReDim HistMensuel(-1 To -1): Exit Function
    
    ' Trier les cles chronologiquement
    Dim keys() As String: keys = dict.Keys()
    Dim i As Integer, j As Integer, tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If keys(i) > keys(j) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i
    
    Dim result() As Double: ReDim result(n-1)
    ReDim periods(n-1)
    For i = 0 To n - 1
        result(i) = CDbl(dict(keys(i)))
        periods(i) = CDate(keys(i) & "-01")
    Next i
    HistMensuel = result
End Function

'--- Classification F2 ---
Private Function ClasserProduit(data() As Double) As String
    Dim n As Integer: n = UBound(data) + 1
    If n < 3 Then ClasserProduit = "INCONNU": Exit Function
    
    Dim avg As Double, i As Integer
    For i = 0 To n-1: avg = avg + data(i): Next i
    avg = avg / n
    If avg < 1 Then ClasserProduit = "VOLATILE": Exit Function
    
    Dim variance As Double
    For i = 0 To n-1: variance = variance + (data(i)-avg)^2: Next i
    Dim cv As Double: cv = Sqr(variance/(n-1)) / avg
    
    ' Pente tendance
    Dim sx As Double, sy As Double, sxy As Double, sx2 As Double
    For i = 0 To n-1
        sx = sx + i: sy = sy + data(i)
        sxy = sxy + i*data(i): sx2 = sx2 + i^2
    Next i
    Dim den As Double: den = n*sx2 - sx^2
    Dim slope As Double
    If Abs(den) > 0.001 Then slope = (n*sxy - sx*sy)/den Else slope = 0
    Dim slopePct As Double: If avg > 0 Then slopePct = Abs(slope)/avg
    
    If cv > CDbl(GetParam("CV_SEUIL_VOLATILE")) Then
        ClasserProduit = "VOLATILE"
    ElseIf slopePct > 0.05 Then
        ClasserProduit = "TENDANCIEL"
    ElseIf cv < CDbl(GetParam("CV_SEUIL_STABLE")) Then
        ClasserProduit = "STABLE"
    Else
        ClasserProduit = "SAISONNIER"
    End If
End Function

'--- Methode 1: Moyenne Mobile 3 ---
Private Sub MoyenneMobile3(data() As Double, horizon As Integer, btLen As Integer, _
    ByRef fc() As Double, ByRef wape As Double, ByRef biais As Double)
    Dim n As Integer: n = UBound(data) + 1
    Dim w As Integer: w = 3
    Dim errS As Double, demS As Double, bS As Double
    Dim i As Integer, j As Integer, pred As Double
    
    For i = w To n - 1
        pred = 0
        For j = i - w To i - 1: pred = pred + data(j): Next j
        pred = pred / w
        If i >= n - btLen Then
            errS = errS + Abs(pred - data(i))
            demS = demS + data(i)
            bS = bS + (pred - data(i))
        End If
    Next i
    
    If demS > 0 Then wape = errS / demS: biais = bS / demS
    
    Dim base As Double: base = 0
    Dim st As Integer: If n >= w Then st = n-w Else st = 0
    For i = st To n - 1: base = base + data(i): Next i
    base = base / Application.Min(w, n)
    ReDim fc(horizon-1)
    For i = 0 To horizon - 1: fc(i) = Application.Max(0, base): Next i
End Sub

'--- Methode 2: Lissage Exponentiel ---
Private Sub LissageExpo(data() As Double, horizon As Integer, btLen As Integer, _
    ByRef fc() As Double, ByRef wape As Double, ByRef biais As Double)
    Dim n As Integer: n = UBound(data) + 1
    Dim alpha As Double: alpha = 0.3
    Dim smooth As Double: smooth = data(0)
    Dim errS As Double, demS As Double, bS As Double, pred As Double
    Dim i As Integer
    
    For i = 1 To n - 1
        pred = smooth
        If i >= n - btLen Then
            errS = errS + Abs(pred - data(i))
            demS = demS + data(i)
            bS = bS + (pred - data(i))
        End If
        smooth = alpha * data(i) + (1 - alpha) * smooth
    Next i
    
    If demS > 0 Then wape = errS / demS: biais = bS / demS
    ReDim fc(horizon-1)
    For i = 0 To horizon - 1: fc(i) = Application.Max(0, smooth): Next i
End Sub

'--- Methode 3: Holt (tendance) ---
Private Sub Holt(data() As Double, horizon As Integer, btLen As Integer, _
    ByRef fc() As Double, ByRef wape As Double, ByRef biais As Double)
    Dim n As Integer: n = UBound(data) + 1
    If n < 2 Then Call LissageExpo(data, horizon, btLen, fc, wape, biais): Exit Sub
    Dim alpha As Double: alpha = 0.3
    Dim beta  As Double: beta  = 0.1
    Dim Lt As Double: Lt = data(0)
    Dim Tt As Double: Tt = data(1) - data(0)
    Dim errS As Double, demS As Double, bS As Double, pred As Double, Lp As Double
    Dim i As Integer
    
    For i = 1 To n - 1
        pred = Application.Max(0, Lt + Tt)
        If i >= n - btLen Then
            errS = errS + Abs(pred - data(i)): demS = demS + data(i): bS = bS + (pred - data(i))
        End If
        Lp = Lt
        Lt = alpha * data(i) + (1 - alpha) * (Lt + Tt)
        Tt = beta * (Lt - Lp) + (1 - beta) * Tt
    Next i
    
    If demS > 0 Then wape = errS / demS: biais = bS / demS
    ReDim fc(horizon-1)
    For i = 1 To horizon: fc(i-1) = Application.Max(0, Lt + i * Tt): Next i
End Sub

'--- Methode 4: Holt-Winters simplifie (saisonnalite 3 periodes) ---
Private Sub HoltWinters(data() As Double, horizon As Integer, btLen As Integer, _
    ByRef fc() As Double, ByRef wape As Double, ByRef biais As Double)
    Dim n As Integer: n = UBound(data) + 1
    Dim seasonLen As Integer: seasonLen = 3
    If n < seasonLen + 2 Then Call Holt(data, horizon, btLen, fc, wape, biais): Exit Sub
    
    Dim alpha As Double: alpha = 0.3
    Dim beta  As Double: beta  = 0.1
    Dim gamma As Double: gamma = 0.2
    Dim i As Integer, j As Integer
    
    Dim initAvg As Double
    For i = 0 To seasonLen - 1: initAvg = initAvg + data(i): Next i
    initAvg = initAvg / seasonLen
    
    Dim SI(2) As Double
    For i = 0 To seasonLen - 1
        If initAvg > 0 Then SI(i) = data(i) / initAvg Else SI(i) = 1
    Next i
    
    Dim Lt As Double: Lt = initAvg
    Dim Tt As Double: Tt = 0
    Dim errS As Double, demS As Double, bS As Double, pred As Double, Lp As Double
    Dim siIdx As Integer
    
    For i = seasonLen To n - 1
        siIdx = i Mod seasonLen
        pred = Application.Max(0, (Lt + Tt) * SI(siIdx))
        If i >= n - btLen Then
            errS = errS + Abs(pred - data(i)): demS = demS + data(i): bS = bS + (pred - data(i))
        End If
        Lp = Lt
        If SI(siIdx) > 0.01 Then Lt = alpha * (data(i) / SI(siIdx)) + (1 - alpha) * (Lt + Tt)
        Tt = beta * (Lt - Lp) + (1 - beta) * Tt
        If Lt > 0.01 Then SI(siIdx) = gamma * (data(i) / Lt) + (1 - gamma) * SI(siIdx)
    Next i
    
    If demS > 0 Then wape = errS / demS: biais = bS / demS
    ReDim fc(horizon-1)
    For i = 0 To horizon - 1
        siIdx = (n + i) Mod seasonLen
        fc(i) = Application.Max(0, (Lt + (i + 1) * Tt) * SI(siIdx))
    Next i
End Sub

' Obtenir la liste des SKU
Private Function GetSKUList() As String()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_PRODUCT)
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then ReDim GetSKUList(-1 To -1): Exit Function
    Dim n As Integer: n = lo.DataBodyRange.Rows.Count
    Dim r() As String: ReDim r(n-1)
    Dim i As Integer
    For i = 1 To n: r(i-1) = CStr(lo.DataBodyRange(i, 1).Value): Next i
    GetSKUList = r
End Function
'@

Set-VBAModule -vbaProject $vbaProject -name "mod03_FORECAST" -code $vba_FORECAST

# ============================================================
# MODULE 04 : PLANS (Demand + Capacity + Inventory + Gap)
# ============================================================
$vba_PLANS = @'
Option Explicit
'==============================================================
' MODULE PLANS - Demand Plan, Capacity Plan, Inventory Plan, Gap Analysis
' F4/F5/F6/F7 â€” EF-04 a EF-07 â€” AtlasFood S&OP DSS
'==============================================================

'-------------------------------------------------------------
' DEMAND PLAN (F4) : agreger le forecast retenu par SKU/mois
'-------------------------------------------------------------
Public Sub LancerDemandPlan()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "DEMANDPLAN", "INFO", "=== DEBUT DEMAND PLAN ==="
    ShowProgress "Demand Plan...", 5
    
    Dim wsFc  As Worksheet: Set wsFc  = ThisWorkbook.Sheets(SH_FORECAST)
    Dim wsDp  As Worksheet: Set wsDp  = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    
    ' Verifier que le forecast existe
    If wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord le Forecast (J2).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ClearSheet SH_DEMAND_PLAN
    
    ' Lire le forecast (methode RETENUE uniquement)
    Dim loFc As ListObject
    On Error Resume Next: Set loFc = wsFc.ListObjects("tbl_FORECAST"): On Error GoTo 0
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cycleID As String: cycleID = GetCycleID()
    
    Dim r As Long
    Dim lastFcRow As Long
    If Not loFc Is Nothing And Not loFc.DataBodyRange Is Nothing Then
        lastFcRow = loFc.DataBodyRange.Rows.Count
    Else
        lastFcRow = wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row - 1
    End If
    
    For r = 2 To lastFcRow + 1
        Dim statut As String
        statut = CStr(wsFc.Cells(r, 7).Value)  ' col 7 = Statut
        If UCase(statut) <> "RETENUE" Then GoTo NextFcRow
        
        Dim sku As String:  sku  = CStr(wsFc.Cells(r, 1).Value)
        Dim pd  As Date:    pd   = wsFc.Cells(r, 2).Value
        Dim qty As Double:  qty  = wsFc.Cells(r, 3).Value
        Dim key As String:  key  = sku & "|" & Format(pd, "YYYY-MM")
        
        If dict.Exists(key) Then dict(key) = dict(key) + qty Else dict.Add key, qty
NextFcRow:
    Next r
    
    ' Ecrire le Demand Plan
    Dim rowOut As Long: rowOut = 2
    Dim k As Variant
    For Each k In dict.Keys()
        Dim parts() As String: parts = Split(CStr(k), "|")
        wsDp.Cells(rowOut, 1).Value = parts(0)  ' SKU_ID
        wsDp.Cells(rowOut, 2).Value = CDate(parts(1) & "-01")  ' Mois
        wsDp.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsDp.Cells(rowOut, 3).Value = Round(CDbl(dict(k)), 0)  ' Qte_stat
        wsDp.Cells(rowOut, 4).Value = 0   ' Ajustement (modifiable par Demand Planner)
        wsDp.Cells(rowOut, 5).Formula = "=C" & rowOut & "+D" & rowOut  ' Qte_finale
        wsDp.Cells(rowOut, 6).Value = cycleID
        wsDp.Cells(rowOut, 7).Value = "Forecast auto"
        rowOut = rowOut + 1
    Next k
    
    ' Creer tableau structure
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsDp.ListObjects("tbl_DEMAND_PLAN"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsDp.ListObjects.Add(xlSrcRange, wsDp.Range("A1:G" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_DEMAND_PLAN": lo.TableStyle = "TableStyleMedium4"
        End If
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    SetStatus "Demand Plan calcule â€” " & Now()
    LogMsg "DEMANDPLAN", "INFO", "Demand Plan OK â€” " & (rowOut-2) & " lignes"
    
    MsgBox "Demand Plan genere." & vbCrLf & _
           "Vous pouvez ajuster la colonne 'Ajustement' avant de continuer." & vbCrLf & _
           "Etape suivante : Capacity + Inventory Plan (J3).", vbInformation, TOOL_NAME
    wsDp.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "DEMANDPLAN", "ERROR", Err.Description
    MsgBox "Erreur Demand Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' CAPACITY PLAN (F5) : charge vs capacite par usine/mois
'-------------------------------------------------------------
Public Sub LancerCapacityPlan()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "CAPACITY", "INFO", "=== DEBUT CAPACITY PLAN ==="
    ShowProgress "Capacity Plan...", 5
    
    Dim wsDp  As Worksheet: Set wsDp  = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsPlant As Worksheet: Set wsPlant = ThisWorkbook.Sheets(SH_PLANT)
    Dim wsProd  As Worksheet: Set wsProd  = ThisWorkbook.Sheets(SH_PRODUCT)
    
    If wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord le Demand Plan (J2).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ClearSheet SH_CAPACITY
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim seuil As Double:   seuil  = CDbl(GetParam("SEUIL_ALERTE_GAP"))
    
    ' Construire: pour chaque usine/mois, somme de la demande des SKU de cette usine
    Dim dictCharge As Object: Set dictCharge = CreateObject("Scripting.Dictionary")
    Dim dictCap    As Object: Set dictCap    = CreateObject("Scripting.Dictionary")
    
    ' Capacites par usine (depuis PLANT)
    Dim loPlant As ListObject
    On Error Resume Next: Set loPlant = wsPlant.ListObjects("tbl_PLANT"): On Error GoTo 0
    Dim plantRow As Long
    If Not loPlant Is Nothing And Not loPlant.DataBodyRange Is Nothing Then
        For plantRow = 1 To loPlant.DataBodyRange.Rows.Count
            Dim plantID As String: plantID = CStr(loPlant.DataBodyRange(plantRow, 1).Value)
            Dim capRef  As Double: capRef  = CDbl(loPlant.DataBodyRange(plantRow, 3).Value)
            Dim marge   As Double: marge   = CDbl(loPlant.DataBodyRange(plantRow, 4).Value)
            Dim capHS   As Double: capHS   = capRef * CDbl(loPlant.DataBodyRange(plantRow, 5).Value)
            Dim capST   As Double: capST   = CDbl(loPlant.DataBodyRange(plantRow, 6).Value)
            ' Cap totale = cap_ref * (1 + marge)
            If Not dictCap.Exists(plantID) Then
                dictCap.Add plantID, Array(capRef * (1 + marge), capHS, capST)
            End If
        Next plantRow
    End If
    
    ' Mapping SKU -> Plant depuis PRODUCT
    Dim skuPlant As Object: Set skuPlant = CreateObject("Scripting.Dictionary")
    Dim loProd As ListObject
    On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
        Dim pr As Long
        For pr = 1 To loProd.DataBodyRange.Rows.Count
            Dim skuID As String: skuID = CStr(loProd.DataBodyRange(pr, 1).Value)
            Dim pID   As String: pID   = CStr(loProd.DataBodyRange(pr, 4).Value)
            If Not skuPlant.Exists(skuID) Then skuPlant.Add skuID, pID
        Next pr
    End If
    
    ' Parcourir le Demand Plan et agregger par usine/mois
    Dim loDp As ListObject
    On Error Resume Next: Set loDp = wsDp.ListObjects("tbl_DEMAND_PLAN"): On Error GoTo 0
    Dim dpRow As Long
    Dim dpLast As Long
    If Not loDp Is Nothing And Not loDp.DataBodyRange Is Nothing Then
        dpLast = loDp.DataBodyRange.Rows.Count
    Else
        dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row - 1
    End If
    
    For dpRow = 1 To dpLast
        Dim dSku  As String: dSku  = CStr(wsDp.Cells(dpRow + 1, 1).Value)
        Dim dDate As Date:   dDate = wsDp.Cells(dpRow + 1, 2).Value
        Dim dQty  As Double
        If IsNumeric(wsDp.Cells(dpRow + 1, 5).Value) Then dQty = CDbl(wsDp.Cells(dpRow + 1, 5).Value)
        
        Dim dPlant As String
        If skuPlant.Exists(dSku) Then dPlant = CStr(skuPlant(dSku)) Else dPlant = "INCONNU"
        
        Dim cKey As String: cKey = dPlant & "|" & Format(dDate, "YYYY-MM")
        If dictCharge.Exists(cKey) Then dictCharge(cKey) = dictCharge(cKey) + dQty _
        Else dictCharge.Add cKey, dQty
    Next dpRow
    
    ' Ecrire Capacity Plan
    Dim rowOut As Long: rowOut = 2
    Dim ck As Variant
    For Each ck In dictCharge.Keys()
        Dim ckParts() As String: ckParts = Split(CStr(ck), "|")
        Dim ckPlant  As String:  ckPlant = ckParts(0)
        Dim ckMois   As Date:    ckMois  = CDate(ckParts(1) & "-01")
        Dim charge   As Double:  charge  = CDbl(dictCharge(ck))
        Dim capN     As Double, capHSv As Double, capSTv As Double
        If dictCap.Exists(ckPlant) Then
            capN   = CDbl(dictCap(ckPlant)(0))
            capHSv = CDbl(dictCap(ckPlant)(1))
            capSTv = CDbl(dictCap(ckPlant)(2))
        Else
            capN = 5000: capHSv = 1000: capSTv = 500
        End If
        
        Dim utilPct As Double
        If capN > 0 Then utilPct = charge / capN Else utilPct = 0
        
        Dim statutCap As String
        Select Case True
            Case utilPct > 1.0:  statutCap = "SURCHARGE"
            Case utilPct > seuil: statutCap = "ALERTE"
            Case Else:             statutCap = "OK"
        End Select
        
        wsCap.Cells(rowOut, 1).Value = ckPlant
        wsCap.Cells(rowOut, 2).Value = ckMois
        wsCap.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsCap.Cells(rowOut, 3).Value = Round(capN, 0)
        wsCap.Cells(rowOut, 4).Value = Round(capHSv, 0)
        wsCap.Cells(rowOut, 5).Value = Round(capSTv, 0)
        wsCap.Cells(rowOut, 6).Value = Round(charge, 0)
        wsCap.Cells(rowOut, 7).Value = Round(utilPct * 100, 1)
        wsCap.Cells(rowOut, 8).Value = statutCap
        wsCap.Cells(rowOut, 9).Value = cycleID
        
        Select Case statutCap
            Case "SURCHARGE": wsCap.Rows(rowOut).Interior.Color = RGB(255, 150, 150)
            Case "ALERTE":    wsCap.Rows(rowOut).Interior.Color = RGB(255, 230, 150)
            Case Else:        wsCap.Rows(rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        rowOut = rowOut + 1
    Next ck
    
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsCap.ListObjects("tbl_CAPACITY"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsCap.ListObjects.Add(xlSrcRange, wsCap.Range("A1:I" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_CAPACITY": lo.TableStyle = "TableStyleMedium5"
        End If
    End If
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Capacity Plan calcule â€” " & Now()
    LogMsg "CAPACITY", "INFO", "Capacity Plan OK â€” " & (rowOut-2) & " lignes"
    
    MsgBox "Capacity Plan genere." & vbCrLf & "Consultez CAPACITY_PLAN.", vbInformation, TOOL_NAME
    wsCap.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "CAPACITY", "ERROR", Err.Description
    MsgBox "Erreur Capacity Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' INVENTORY PLAN (F6) : projection stock + detection rupture
'-------------------------------------------------------------
Public Sub LancerInventoryPlan()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "INVENTORY", "INFO", "=== DEBUT INVENTORY PLAN ==="
    ShowProgress "Inventory Plan...", 5
    
    Dim wsDp  As Worksheet: Set wsDp  = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets(SH_INVENTORY)
    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    
    If wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord le Demand Plan (J2).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ClearSheet SH_INVENTORY
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim nivServ As Double: nivServ = CDbl(GetParam("NIVEAU_SERVICE"))
    Dim delai   As Double: delai   = CDbl(GetParam("DELAI_LIVRAISON"))
    Dim seuil2  As Double: seuil2  = CDbl(GetParam("SEUIL_SURSTOCK"))
    
    ' Pour chaque SKU, parcourir les mois du Demand Plan
    Dim skuList As Object: Set skuList = CreateObject("Scripting.Dictionary")
    
    ' Construire dictionnaire SKU -> dict(mois -> qte_finale)
    Dim loDp As ListObject
    On Error Resume Next: Set loDp = wsDp.ListObjects("tbl_DEMAND_PLAN"): On Error GoTo 0
    Dim dpLast As Long
    If Not loDp Is Nothing And Not loDp.DataBodyRange Is Nothing Then
        dpLast = loDp.DataBodyRange.Rows.Count
    Else
        dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row - 1
    End If
    
    Dim dpRow As Long
    For dpRow = 1 To dpLast
        Dim dSku As String: dSku = CStr(wsDp.Cells(dpRow+1, 1).Value)
        Dim dMois As Date:  dMois = wsDp.Cells(dpRow+1, 2).Value
        Dim dQty  As Double
        If IsNumeric(wsDp.Cells(dpRow+1, 5).Value) Then dQty = CDbl(wsDp.Cells(dpRow+1, 5).Value)
        Dim dKey As String: dKey = dSku & "|" & Format(dMois, "YYYY-MM")
        If Not skuList.Exists(dKey) Then skuList.Add dKey, dQty
    Next dpRow
    
    ' Mapping DLC depuis PRODUCT
    Dim dlcMap As Object: Set dlcMap = CreateObject("Scripting.Dictionary")
    Dim loProd As ListObject
    On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
        Dim pr As Long
        For pr = 1 To loProd.DataBodyRange.Rows.Count
            Dim sID  As String: sID  = CStr(loProd.DataBodyRange(pr, 1).Value)
            Dim dlcJ As Integer
            If IsNumeric(loProd.DataBodyRange(pr, 6).Value) Then dlcJ = CInt(loProd.DataBodyRange(pr, 6).Value)
            If Not dlcMap.Exists(sID) Then dlcMap.Add sID, dlcJ
        Next pr
    End If
    
    ' Stocks initiaux depuis PARAMETRES
    Dim stockInit As Object: Set stockInit = CreateObject("Scripting.Dictionary")
    stockInit("SKU-01") = CDbl(GetParam("STOCK_INIT_SKU01"))
    stockInit("SKU-02") = CDbl(GetParam("STOCK_INIT_SKU02"))
    stockInit("SKU-03") = CDbl(GetParam("STOCK_INIT_SKU03"))
    stockInit("SKU-04") = CDbl(GetParam("STOCK_INIT_SKU04"))
    stockInit("SKU-05") = CDbl(GetParam("STOCK_INIT_SKU05"))
    
    ' Trier les cles par SKU puis par mois
    Dim allKeys() As String: allKeys = skuList.Keys()
    Dim nk As Integer: nk = skuList.Count
    Dim ii As Integer, jj As Integer, tmpK As String
    For ii = 0 To nk - 2
        For jj = ii + 1 To nk - 1
            If allKeys(ii) > allKeys(jj) Then tmpK = allKeys(ii): allKeys(ii) = allKeys(jj): allKeys(jj) = tmpK
        Next jj
    Next ii
    
    ' Suivre stock courant par SKU
    Dim currStock As Object: Set currStock = CreateObject("Scripting.Dictionary")
    
    Dim rowOut As Long: rowOut = 2
    For ii = 0 To nk - 1
        Dim k As String: k = allKeys(ii)
        Dim parts() As String: parts = Split(k, "|")
        Dim iSku  As String: iSku  = parts(0)
        Dim iMois As Date:   iMois = CDate(parts(1) & "-01")
        Dim demand As Double: demand = CDbl(skuList(k))
        
        ' Stock debut
        Dim stockDeb As Double
        If currStock.Exists(iSku) Then
            stockDeb = CDbl(currStock(iSku))
        Else
            If stockInit.Exists(iSku) Then stockDeb = CDbl(stockInit(iSku)) Else stockDeb = 1000
        End If
        
        ' Stock de securite (formule BR-21: SS = Z * sigma * sqrt(delai))
        ' Simplification: SS = demande_moy * (delai/30) * Z_service
        Dim Z As Double
        Select Case True
            Case nivServ >= 0.99: Z = 2.33
            Case nivServ >= 0.98: Z = 2.05
            Case nivServ >= 0.95: Z = 1.65
            Case nivServ >= 0.90: Z = 1.28
            Case Else:             Z = 1.0
        End Select
        Dim cv As Double: cv = 0.2  ' Approximation CV
        Dim stockSecu As Double: stockSecu = Round(Z * cv * demand * Sqr(delai / 30), 0)
        
        ' Production = demande (equilibre par defaut dans MVP)
        ' En scenario, on pourra ajouter HS/ST
        Dim prod  As Double: prod  = demand
        Dim livr  As Double: livr  = demand
        Dim stockFin As Double: stockFin = stockDeb + prod - livr
        
        ' DLC restante estimee (jours)
        Dim dlcRest As Double: dlcRest = 0
        If dlcMap.Exists(iSku) Then
            dlcRest = dlcMap(iSku) - 30  ' 1 mois ecoule
        End If
        
        ' Statut stock
        Dim statutInv As String
        Select Case True
            Case stockFin < 0:                  statutInv = "RUPTURE"
            Case stockFin < stockSecu:           statutInv = "SOUS_SEUIL"
            Case stockFin > demand * seuil2:     statutInv = "SURSTOCK"
            Case Else:                            statutInv = "OK"
        End Select
        
        wsInv.Cells(rowOut, 1).Value = iSku
        wsInv.Cells(rowOut, 2).Value = iMois
        wsInv.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsInv.Cells(rowOut, 3).Value = Round(stockDeb, 0)
        wsInv.Cells(rowOut, 4).Value = Round(stockSecu, 0)
        wsInv.Cells(rowOut, 5).Value = Round(prod, 0)
        wsInv.Cells(rowOut, 6).Value = Round(livr, 0)
        wsInv.Cells(rowOut, 7).Value = Round(stockFin, 0)
        wsInv.Cells(rowOut, 8).Value = Round(dlcRest, 0)
        wsInv.Cells(rowOut, 9).Value = statutInv
        wsInv.Cells(rowOut, 10).Value = cycleID
        
        Select Case statutInv
            Case "RUPTURE":    wsInv.Rows(rowOut).Interior.Color = RGB(255, 100, 100)
            Case "SOUS_SEUIL": wsInv.Rows(rowOut).Interior.Color = RGB(255, 200, 150)
            Case "SURSTOCK":   wsInv.Rows(rowOut).Interior.Color = RGB(200, 200, 255)
            Case Else:          wsInv.Rows(rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        
        ' Mettre a jour le stock courant
        If currStock.Exists(iSku) Then currStock(iSku) = stockFin Else currStock.Add iSku, stockFin
        rowOut = rowOut + 1
    Next ii
    
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsInv.ListObjects("tbl_INVENTORY"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsInv.ListObjects.Add(xlSrcRange, wsInv.Range("A1:J" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_INVENTORY": lo.TableStyle = "TableStyleMedium6"
        End If
    End If
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Inventory Plan calcule â€” " & Now()
    LogMsg "INVENTORY", "INFO", "Inventory Plan OK â€” " & (rowOut-2) & " lignes"
    
    MsgBox "Inventory Plan genere." & vbCrLf & "Etape suivante: Gap Analysis (J3).", vbInformation, TOOL_NAME
    wsInv.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "INVENTORY", "ERROR", Err.Description
    MsgBox "Erreur Inventory Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' GAP ANALYSIS (F7) : Besoin - Capacite par usine/mois
'-------------------------------------------------------------
Public Sub LancerGapAnalysis()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "GAP", "INFO", "=== DEBUT GAP ANALYSIS ==="
    ShowProgress "Gap Analysis...", 5
    
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsGap As Worksheet: Set wsGap = ThisWorkbook.Sheets(SH_GAP)
    
    If wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord le Capacity Plan (J3).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ClearSheet SH_GAP
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim seuil   As Double: seuil   = CDbl(GetParam("SEUIL_ALERTE_GAP"))
    
    ' Lire le Capacity Plan
    Dim loCap As ListObject
    On Error Resume Next: Set loCap = wsCap.ListObjects("tbl_CAPACITY"): On Error GoTo 0
    Dim capLast As Long
    If Not loCap Is Nothing And Not loCap.DataBodyRange Is Nothing Then
        capLast = loCap.DataBodyRange.Rows.Count
    Else
        capLast = wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row - 1
    End If
    
    Dim rowOut As Long: rowOut = 2
    Dim cr As Long
    For cr = 1 To capLast
        Dim capPlant  As String: capPlant  = CStr(wsCap.Cells(cr+1, 1).Value)
        Dim capMois   As Date:   capMois   = wsCap.Cells(cr+1, 2).Value
        Dim capNorm   As Double: capNorm   = CDbl(wsCap.Cells(cr+1, 3).Value)
        Dim capHSm    As Double: capHSm    = CDbl(wsCap.Cells(cr+1, 4).Value)
        Dim capSTm    As Double: capSTm    = CDbl(wsCap.Cells(cr+1, 5).Value)
        Dim chargeV   As Double: chargeV   = CDbl(wsCap.Cells(cr+1, 6).Value)
        
        Dim capDispo  As Double: capDispo  = capNorm + capHSm + capSTm
        Dim gap       As Double: gap       = capDispo - chargeV
        Dim gapPct    As Double: If capDispo > 0 Then gapPct = gap / capDispo Else gapPct = 0
        
        Dim statutGap As String
        Select Case True
            Case gap < 0:             statutGap = "DEFICIT"
            Case chargeV / capNorm > seuil: statutGap = "TENSION"
            Case Else:                 statutGap = "OK"
        End Select
        
        wsGap.Cells(rowOut, 1).Value = capPlant
        wsGap.Cells(rowOut, 2).Value = capMois
        wsGap.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsGap.Cells(rowOut, 3).Value = Round(chargeV, 0)
        wsGap.Cells(rowOut, 4).Value = Round(capDispo, 0)
        wsGap.Cells(rowOut, 5).Value = Round(gap, 0)
        wsGap.Cells(rowOut, 6).Value = Round(gapPct * 100, 1)
        wsGap.Cells(rowOut, 7).Value = statutGap
        wsGap.Cells(rowOut, 8).Value = cycleID
        
        Select Case statutGap
            Case "DEFICIT": wsGap.Rows(rowOut).Interior.Color = RGB(255, 100, 100)
            Case "TENSION": wsGap.Rows(rowOut).Interior.Color = RGB(255, 220, 120)
            Case Else:       wsGap.Rows(rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        rowOut = rowOut + 1
    Next cr
    
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsGap.ListObjects("tbl_GAP"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsGap.ListObjects.Add(xlSrcRange, wsGap.Range("A1:H" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_GAP": lo.TableStyle = "TableStyleMedium3"
        End If
    End If
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Gap Analysis calcule â€” " & Now()
    LogMsg "GAP", "INFO", "Gap Analysis OK â€” " & (rowOut-2) & " lignes"
    
    MsgBox "Gap Analysis terminee." & vbCrLf & "Etape suivante: Scenarios (J4).", vbInformation, TOOL_NAME
    wsGap.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "GAP", "ERROR", Err.Description
    MsgBox "Erreur Gap Analysis: " & Err.Description, vbCritical, TOOL_NAME
End Sub
'@

Set-VBAModule -vbaProject $vbaProject -name "mod04_PLANS" -code $vba_PLANS

# ============================================================
# MODULE 05 : SCENARIOS
# ============================================================
$vba_SCENARIOS = @'
Option Explicit
'==============================================================
' MODULE SCENARIOS - Gestionnaire de 5 scenarios S&OP
' F10 / EF-14 â€” AtlasFood S&OP DSS (niveau Avance)
'==============================================================

Public Sub GenererScenarios()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "SCENARIOS", "INFO", "=== DEBUT GENERATION SCENARIOS ==="
    ShowProgress "Generation des scenarios...", 5
    
    ' Verifier Gap Analysis
    Dim wsGap As Worksheet: Set wsGap = ThisWorkbook.Sheets(SH_GAP)
    If wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord la Gap Analysis (J3).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Sheets(SH_SCENARIOS)
    ClearSheet SH_SCENARIOS
    
    Dim cycleID As String: cycleID = GetCycleID()
    
    ' Lire les KPI du Gap Analysis (agregation globale)
    Dim loGap As ListObject
    On Error Resume Next: Set loGap = wsGap.ListObjects("tbl_GAP"): On Error GoTo 0
    Dim gapRows As Long
    If Not loGap Is Nothing And Not loGap.DataBodyRange Is Nothing Then
        gapRows = loGap.DataBodyRange.Rows.Count
    Else
        gapRows = wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row - 1
    End If
    
    ' Calcul agregat (charge totale, capacite totale, deficit total)
    Dim totCharge As Double, totCap As Double, totGap As Double
    Dim nDeficit As Integer
    Dim gr As Long
    For gr = 1 To gapRows
        Dim gCharge  As Double: If IsNumeric(wsGap.Cells(gr+1,3).Value) Then gCharge = CDbl(wsGap.Cells(gr+1,3).Value)
        Dim gCapD    As Double: If IsNumeric(wsGap.Cells(gr+1,4).Value) Then gCapD   = CDbl(wsGap.Cells(gr+1,4).Value)
        Dim gGap     As Double: If IsNumeric(wsGap.Cells(gr+1,5).Value) Then gGap    = CDbl(wsGap.Cells(gr+1,5).Value)
        Dim gStatut  As String: gStatut = CStr(wsGap.Cells(gr+1,7).Value)
        totCharge = totCharge + gCharge
        totCap    = totCap    + gCapD
        totGap    = totGap    + gGap
        If gStatut = "DEFICIT" Then nDeficit = nDeficit + 1
    Next gr
    
    ' Lire les couts depuis COST
    Dim wsCost As Worksheet: Set wsCost = ThisWorkbook.Sheets(SH_COST)
    Dim coutHS As Double, coutST As Double, coutStock As Double, coutPen As Double, coutProd As Double
    coutHS    = GetCoutMoyen("HS",         wsCost)
    coutST    = GetCoutMoyen("ST",         wsCost)
    coutStock = GetCoutMoyen("Stockage",   wsCost)
    coutPen   = GetCoutMoyen("Penurie",    wsCost)
    coutProd  = GetCoutMoyen("Prod",       wsCost)
    
    ' Demande totale (horizon)
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim demTot As Double: demTot = Application.Sum(wsDp.Columns(5))
    If demTot <= 0 Then demTot = 50000  ' Valeur par defaut
    
    ' Gap negatif total (deficit de capacite)
    Dim deficitAbs As Double: deficitAbs = Abs(Application.Min(0, totGap))
    
    '------------------------------------------------------
    ' 5 SCENARIOS
    '------------------------------------------------------
    Dim scenarios(4, 9) As Variant
    
    ' SC1 - BASELINE : rien ne change
    scenarios(0,0) = "SC-01"
    scenarios(0,1) = "Baseline (Statu quo)"
    scenarios(0,2) = "Aucun levier active. Production a capacite normale."
    Dim coutSC1 As Double: coutSC1 = demTot * coutProd + deficitAbs * coutPen
    scenarios(0,3) = Round(coutSC1, 0)
    Dim svcSC1 As Double: If totCharge > 0 Then svcSC1 = (totCharge - deficitAbs) / totCharge Else svcSC1 = 1
    scenarios(0,4) = Round(Application.Max(0, Application.Min(1, svcSC1)) * 100, 1)
    scenarios(0,5) = Round(deficitAbs * 0.1, 0)  ' Stock final approx
    scenarios(0,6) = Round(deficitAbs * 0.05, 0)  ' Risque peremption
    If totCap > 0 Then scenarios(0,7) = Round(totCharge / totCap * 100, 1) Else scenarios(0,7) = 100
    scenarios(0,8) = cycleID
    scenarios(0,9) = "Brouillon"
    
    ' SC2 - HEURES SUP
    scenarios(1,0) = "SC-02"
    scenarios(1,1) = "Heures supplementaires (+20%)"
    scenarios(1,2) = "Activation HS sur usines en deficit. Cap +20%."
    Dim defHS As Double: defHS = Application.Max(0, deficitAbs - totCap * 0.2)
    Dim coutSC2 As Double: coutSC2 = demTot * coutProd + totCap * 0.2 * coutHS + defHS * coutPen
    scenarios(1,3) = Round(coutSC2, 0)
    Dim svcSC2 As Double: If totCharge > 0 Then svcSC2 = (totCharge - defHS) / totCharge Else svcSC2 = 1
    scenarios(1,4) = Round(Application.Max(0, Application.Min(1, svcSC2)) * 100, 1)
    scenarios(1,5) = Round(defHS * 0.1, 0)
    scenarios(1,6) = Round(totCap * 0.2 * 0.03, 0)
    If totCap > 0 Then scenarios(1,7) = Round((totCharge) / (totCap * 1.2) * 100, 1) Else scenarios(1,7) = 90
    scenarios(1,8) = cycleID
    scenarios(1,9) = "Brouillon"
    
    ' SC3 - SOUS-TRAITANCE
    scenarios(2,0) = "SC-03"
    scenarios(2,1) = "Sous-traitance partielle"
    scenarios(2,2) = "Externalisation 50% du deficit. Delai +5 jours."
    Dim defST As Double: defST = Application.Max(0, deficitAbs - deficitAbs * 0.5)
    Dim coutSC3 As Double: coutSC3 = demTot * coutProd + deficitAbs * 0.5 * coutST + defST * coutPen
    scenarios(2,3) = Round(coutSC3, 0)
    Dim svcSC3 As Double: If totCharge > 0 Then svcSC3 = (totCharge - defST) / totCharge Else svcSC3 = 1
    scenarios(2,4) = Round(Application.Max(0, Application.Min(1, svcSC3)) * 100, 1)
    scenarios(2,5) = Round(defST * 0.1, 0)
    scenarios(2,6) = Round(deficitAbs * 0.5 * 0.03, 0)
    If totCap > 0 Then scenarios(2,7) = Round(totCharge / totCap * 100, 1) Else scenarios(2,7) = 100
    scenarios(2,8) = cycleID
    scenarios(2,9) = "Brouillon"
    
    ' SC4 - HS + ST
    scenarios(3,0) = "SC-04"
    scenarios(3,1) = "Heures sup + Sous-traitance"
    scenarios(3,2) = "HS 20% + ST pour deficit residuel. Solution optimale cout/service."
    Dim capTotSC4 As Double: capTotSC4 = totCap * 1.2 + deficitAbs * 0.5
    Dim defSC4 As Double: defSC4 = Application.Max(0, deficitAbs - (totCap * 0.2 + deficitAbs * 0.5))
    Dim coutSC4 As Double: coutSC4 = demTot*coutProd + totCap*0.2*coutHS + deficitAbs*0.5*coutST + defSC4*coutPen
    scenarios(3,3) = Round(coutSC4, 0)
    Dim svcSC4 As Double: If totCharge > 0 Then svcSC4 = (totCharge - defSC4) / totCharge Else svcSC4 = 1
    scenarios(3,4) = Round(Application.Max(0, Application.Min(1, svcSC4)) * 100, 1)
    scenarios(3,5) = Round(defSC4 * 0.1, 0)
    scenarios(3,6) = Round((totCap*0.2 + deficitAbs*0.5) * 0.03, 0)
    If capTotSC4 > 0 Then scenarios(3,7) = Round(totCharge / capTotSC4 * 100, 1) Else scenarios(3,7) = 80
    scenarios(3,8) = cycleID
    scenarios(3,9) = "Brouillon"
    
    ' SC5 - LISSAGE DEMANDE
    scenarios(4,0) = "SC-05"
    scenarios(4,1) = "Lissage de la demande (-10%)"
    scenarios(4,2) = "Revision a la baisse de 10% de la demande (promotions reportees)."
    Dim demLisse As Double: demLisse = demTot * 0.9
    Dim defLisse As Double: defLisse = Application.Max(0, deficitAbs - demTot * 0.1)
    Dim coutSC5 As Double: coutSC5 = demLisse * coutProd + defLisse * coutPen + (demTot - demLisse) * coutStock
    scenarios(4,3) = Round(coutSC5, 0)
    Dim svcSC5 As Double: If totCharge > 0 Then svcSC5 = (demLisse - defLisse) / totCharge Else svcSC5 = 0.9
    scenarios(4,4) = Round(Application.Max(0, Application.Min(1, svcSC5)) * 100, 1)
    scenarios(4,5) = Round(defLisse * 0.1, 0)
    scenarios(4,6) = Round(demTot * 0.1 * 0.02, 0)
    If totCap > 0 Then scenarios(4,7) = Round(demLisse / totCap * 100, 1) Else scenarios(4,7) = 90
    scenarios(4,8) = cycleID
    scenarios(4,9) = "Brouillon"
    
    ' Ecrire les scenarios
    Dim rowOut As Long: rowOut = 2
    Dim s As Integer
    For s = 0 To 4
        For Dim c As Integer = 0 To 9
            wsSc.Cells(rowOut, c+1).Value = scenarios(s, c)
        Next c
        ' Colorisation
        Select Case s
            Case 0: wsSc.Rows(rowOut).Interior.Color = RGB(220, 220, 220)
            Case 1: wsSc.Rows(rowOut).Interior.Color = RGB(200, 230, 255)
            Case 2: wsSc.Rows(rowOut).Interior.Color = RGB(255, 230, 200)
            Case 3: wsSc.Rows(rowOut).Interior.Color = RGB(200, 255, 200)
            Case 4: wsSc.Rows(rowOut).Interior.Color = RGB(255, 255, 200)
        End Select
        rowOut = rowOut + 1
    Next s
    
    ' Creer tableau
    Dim lo As ListObject
    On Error Resume Next: Set lo = wsSc.ListObjects("tbl_SCENARIOS"): On Error GoTo 0
    If lo Is Nothing Then
        Set lo = wsSc.ListObjects.Add(xlSrcRange, wsSc.Range("A1:J6"), , xlYes)
        lo.Name = "tbl_SCENARIOS": lo.TableStyle = "TableStyleMedium7"
    End If
    
    ' Ajuster largeur
    wsSc.Columns("C").ColumnWidth = 45
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "5 scenarios generes â€” " & Now()
    LogMsg "SCENARIOS", "INFO", "5 scenarios generes."
    
    MsgBox "5 scenarios generes dans SCENARIOS." & vbCrLf & _
           "Comparez les KPI (cout, service, utilisation)." & vbCrLf & _
           "Utilisez 'Valider Plan' pour retenir un scenario.", vbInformation, TOOL_NAME
    wsSc.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "SCENARIOS", "ERROR", Err.Description
    MsgBox "Erreur Scenarios: " & Err.Description, vbCritical, TOOL_NAME
End Sub

' Calculer le cout moyen d'un type pour toutes les usines
Private Function GetCoutMoyen(typeStr As String, wsCost As Worksheet) As Double
    Dim lo As ListObject
    On Error Resume Next: Set lo = wsCost.ListObjects("tbl_COST"): On Error GoTo 0
    Dim total As Double: Dim count As Integer: count = 0
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.Count
            If CStr(lo.DataBodyRange(r, 1).Value) = typeStr Then
                If IsNumeric(lo.DataBodyRange(r, 3).Value) Then
                    total = total + CDbl(lo.DataBodyRange(r, 3).Value)
                    count = count + 1
                End If
            End If
        Next r
    End If
    If count > 0 Then GetCoutMoyen = total / count Else GetCoutMoyen = 200
End Function
'@

Set-VBAModule -vbaProject $vbaProject -name "mod05_SCENARIOS" -code $vba_SCENARIOS

# ============================================================
# MODULE 06 : KPI + PLAN FINAL + MONTE CARLO
# ============================================================
$vba_KPI = @'
Option Explicit
'==============================================================
' MODULE KPI - Calcul KPI, Plan Final, Monte Carlo
' F25/F23/F18 â€” EF-08/EF-13/EF-23
'==============================================================

'-------------------------------------------------------------
' CALCUL KPI (F25 / EF-08)
'-------------------------------------------------------------
Public Sub CalculerKPI()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "KPI", "INFO", "=== CALCUL KPI ==="
    ShowProgress "Calcul KPI...", 5
    
    Dim wsKpi As Worksheet: Set wsKpi = ThisWorkbook.Sheets(SH_KPI)
    Dim wsGap As Worksheet: Set wsGap = ThisWorkbook.Sheets(SH_GAP)
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets(SH_INVENTORY)
    Dim wsFc  As Worksheet: Set wsFc  = ThisWorkbook.Sheets(SH_FORECAST)
    Dim wsDp  As Worksheet: Set wsDp  = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    
    ClearSheet SH_KPI
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim rowOut As Long: rowOut = 2
    Dim kpiIdx As Long: kpiIdx = 1
    
    ' --- KPI 1 : Taux de service global ---
    Dim totDem As Double: totDem = Application.Sum(wsDp.Columns(5))
    Dim totRup As Double
    Dim invLast As Long: invLast = wsInv.Cells(wsInv.Rows.Count,1).End(xlUp).Row
    Dim ir As Long
    For ir = 2 To invLast
        Dim invSt As String: invSt = CStr(wsInv.Cells(ir, 9).Value)
        If invSt = "RUPTURE" Then
            If IsNumeric(wsInv.Cells(ir, 6).Value) Then totRup = totRup + CDbl(wsInv.Cells(ir,6).Value)
        End If
    Next ir
    Dim tauxService As Double
    If totDem > 0 Then tauxService = (totDem - totRup) / totDem Else tauxService = 1
    AjouterKPI wsKpi, rowOut, kpiIdx, "Taux de service", "Global", "Horizon " & GetHorizon() & " mois", _
        Round(tauxService * 100, 1), "%", 95, 90, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 2 : Utilisation max capacite ---
    Dim maxUtil As Double: maxUtil = 0
    Dim capLast As Long: capLast = wsCap.Cells(wsCap.Rows.Count,1).End(xlUp).Row
    Dim cr As Long
    For cr = 2 To capLast
        If IsNumeric(wsCap.Cells(cr, 7).Value) Then
            If CDbl(wsCap.Cells(cr, 7).Value) > maxUtil Then maxUtil = CDbl(wsCap.Cells(cr, 7).Value)
        End If
    Next cr
    AjouterKPI wsKpi, rowOut, kpiIdx, "Utilisation max capacite", "Global", "Peak horizon", _
        Round(maxUtil, 1), "%", 85, 95, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 3 : WAPE moyen (methode retenue) ---
    Dim sumWAPE As Double: Dim nWAPE As Long: nWAPE = 0
    Dim fcLast As Long: fcLast = wsFc.Cells(wsFc.Rows.Count,1).End(xlUp).Row
    Dim fr As Long
    For fr = 2 To fcLast
        If CStr(wsFc.Cells(fr, 7).Value) = "RETENUE" Then
            If IsNumeric(wsFc.Cells(fr, 5).Value) Then
                sumWAPE = sumWAPE + CDbl(wsFc.Cells(fr, 5).Value)
                nWAPE = nWAPE + 1
            End If
        End If
    Next fr
    Dim wapeMoy As Double: If nWAPE > 0 Then wapeMoy = sumWAPE / nWAPE
    AjouterKPI wsKpi, rowOut, kpiIdx, "WAPE moyen forecast", "Global", "Methode retenue", _
        Round(wapeMoy, 1), "%", 20, 30, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 4 : Nombre de ruptures ---
    Dim nRuptures As Long: nRuptures = 0
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then nRuptures = nRuptures + 1
    Next ir
    AjouterKPI wsKpi, rowOut, kpiIdx, "Nombre de ruptures", "Inventory", "Horizon", _
        nRuptures, "cas", 0, 3, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 5 : Stock total final ---
    Dim totStockFin As Double
    For ir = 2 To invLast
        If IsNumeric(wsInv.Cells(ir, 7).Value) Then totStockFin = totStockFin + CDbl(wsInv.Cells(ir,7).Value)
    Next ir
    AjouterKPI wsKpi, rowOut, kpiIdx, "Stock total fin horizon", "Inventory", "Horizon", _
        Round(totStockFin, 0), "unites", 0, 0, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 6 : Nombre usines en surcharge ---
    Dim nSurcharge As Long: nSurcharge = 0
    For cr = 2 To capLast
        If CStr(wsCap.Cells(cr, 8).Value) = "SURCHARGE" Then nSurcharge = nSurcharge + 1
    Next cr
    AjouterKPI wsKpi, rowOut, kpiIdx, "Usines/mois en surcharge", "Capacity", "Horizon", _
        nSurcharge, "cas", 0, 2, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' --- KPI 7 : Deficit capacite total ---
    Dim totDeficit As Double
    Dim gapLast As Long: gapLast = wsGap.Cells(wsGap.Rows.Count,1).End(xlUp).Row
    Dim gr As Long
    For gr = 2 To gapLast
        Dim gVal As Double
        If IsNumeric(wsGap.Cells(gr, 5).Value) Then gVal = CDbl(wsGap.Cells(gr, 5).Value)
        If gVal < 0 Then totDeficit = totDeficit + Abs(gVal)
    Next gr
    AjouterKPI wsKpi, rowOut, kpiIdx, "Deficit capacite total", "Gap", "Horizon", _
        Round(totDeficit, 0), "unites", 0, 1, cycleID
    rowOut = rowOut + 1: kpiIdx = kpiIdx + 1
    
    ' Creer tableau
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsKpi.ListObjects("tbl_KPI"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsKpi.ListObjects.Add(xlSrcRange, wsKpi.Range("A1:J" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_KPI": lo.TableStyle = "TableStyleMedium2"
        End If
        wsKpi.Columns("B").ColumnWidth = 28
        wsKpi.Columns("E").ColumnWidth = 12
    End If
    
    ' Mettre a jour les alertes
    Call GenererAlertes
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "KPI calcules â€” " & Now()
    LogMsg "KPI", "INFO", "KPI calcules."
    
    MsgBox "KPI calcules dans KPI_RESULTS." & vbCrLf & _
           "Alertes generees dans ALERTES." & vbCrLf & vbCrLf & _
           "Etape suivante: Valider le plan (J5).", vbInformation, TOOL_NAME
    wsKpi.Activate
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "KPI", "ERROR", Err.Description
    MsgBox "Erreur KPI: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Private Sub AjouterKPI(ws As Worksheet, rowOut As Long, kpiIdx As Long, _
    indicateur As String, perimetre As String, periode As String, _
    valeur As Double, unite As String, seuilVert As Double, seuilRouge As Double, _
    cycleID As String)
    
    ws.Cells(rowOut, 1).Value = "KPI-" & Format(kpiIdx, "00")
    ws.Cells(rowOut, 2).Value = indicateur
    ws.Cells(rowOut, 3).Value = perimetre
    ws.Cells(rowOut, 4).Value = periode
    ws.Cells(rowOut, 5).Value = valeur
    ws.Cells(rowOut, 6).Value = unite
    ws.Cells(rowOut, 7).Value = seuilVert
    ws.Cells(rowOut, 8).Value = seuilRouge
    ws.Cells(rowOut, 10).Value = cycleID
    
    ' Statut seuil
    Dim statut As String: statut = "OK"
    Select Case indicateur
        Case "Taux de service", "WAPE moyen forecast"  ' Plus = mieux pour service, moins = mieux pour WAPE
            If indicateur = "Taux de service" Then
                If valeur >= seuilVert Then statut = "VERT" ElseIf valeur >= seuilRouge Then statut = "ORANGE" Else statut = "ROUGE"
            Else
                If valeur <= seuilVert Then statut = "VERT" ElseIf valeur <= seuilRouge Then statut = "ORANGE" Else statut = "ROUGE"
            End If
        Case Else  ' Moins = mieux
            If valeur <= seuilVert Then statut = "VERT" ElseIf valeur <= seuilRouge Then statut = "ORANGE" Else statut = "ROUGE"
    End Select
    ws.Cells(rowOut, 9).Value = statut
    
    Select Case statut
        Case "VERT":   ws.Cells(rowOut, 9).Interior.Color = RGB(150, 255, 150)
        Case "ORANGE": ws.Cells(rowOut, 9).Interior.Color = RGB(255, 200, 100)
        Case "ROUGE":  ws.Cells(rowOut, 9).Interior.Color = RGB(255, 100, 100)
    End Select
End Sub

'-------------------------------------------------------------
' ALERTES
'-------------------------------------------------------------
Private Sub GenererAlertes()
    Dim wsAlt As Worksheet: Set wsAlt = ThisWorkbook.Sheets(SH_ALERTS)
    ClearSheet SH_ALERTS
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim rowOut As Long: rowOut = 2
    Dim alertID As Long: alertID = 1
    
    ' Ruptures de stock
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets(SH_INVENTORY)
    Dim ir As Long
    For ir = 2 To wsInv.Cells(wsInv.Rows.Count,1).End(xlUp).Row
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then
            wsAlt.Cells(rowOut, 1).Value = "ALT-" & Format(alertID, "000")
            wsAlt.Cells(rowOut, 2).Value = "RUPTURE_STOCK"
            wsAlt.Cells(rowOut, 3).Value = "CRITIQUE"
            wsAlt.Cells(rowOut, 4).Value = wsInv.Cells(ir, 1).Value & " / " & wsInv.Cells(ir, 2).Text
            wsAlt.Cells(rowOut, 5).Value = "Rupture detectee: stock fin negatif"
            wsAlt.Cells(rowOut, 6).Value = cycleID
            wsAlt.Cells(rowOut, 7).Value = Now()
            wsAlt.Rows(rowOut).Interior.Color = RGB(255, 150, 150)
            rowOut = rowOut + 1: alertID = alertID + 1
        End If
    Next ir
    
    ' Surcharges capacite
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim cr As Long
    For cr = 2 To wsCap.Cells(wsCap.Rows.Count,1).End(xlUp).Row
        If CStr(wsCap.Cells(cr, 8).Value) = "SURCHARGE" Then
            wsAlt.Cells(rowOut, 1).Value = "ALT-" & Format(alertID, "000")
            wsAlt.Cells(rowOut, 2).Value = "SURCHARGE_CAPACITE"
            wsAlt.Cells(rowOut, 3).Value = "ELEVEE"
            wsAlt.Cells(rowOut, 4).Value = wsCap.Cells(cr, 1).Value & " / " & wsCap.Cells(cr, 2).Text
            wsAlt.Cells(rowOut, 5).Value = "Surcharge: utilisation=" & wsCap.Cells(cr, 7).Value & "%"
            wsAlt.Cells(rowOut, 6).Value = cycleID
            wsAlt.Cells(rowOut, 7).Value = Now()
            wsAlt.Rows(rowOut).Interior.Color = RGB(255, 220, 150)
            rowOut = rowOut + 1: alertID = alertID + 1
        End If
    Next cr
    
    LogMsg "KPI", "INFO", (alertID - 1) & " alerte(s) generee(s)"
End Sub

'-------------------------------------------------------------
' VALIDER LE PLAN FINAL (F23 / EF-13)
'-------------------------------------------------------------
Public Sub ValiderPlanFinal()
    On Error GoTo ErrHandler
    
    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Sheets(SH_SCENARIOS)
    If wsSc.Cells(wsSc.Rows.Count,1).End(xlUp).Row < 2 Then
        MsgBox "Generez d'abord les scenarios (J4).", vbCritical, TOOL_NAME
        Exit Sub
    End If
    
    ' Afficher dialogue de choix de scenario
    Dim scenStr As String
    scenStr = "Choisissez le scenario a valider:" & vbCrLf & vbCrLf
    Dim sr As Long
    For sr = 2 To wsSc.Cells(wsSc.Rows.Count,1).End(xlUp).Row
        scenStr = scenStr & wsSc.Cells(sr, 1).Value & " - " & wsSc.Cells(sr, 2).Value & vbCrLf
        scenStr = scenStr & "   Cout: " & Format(wsSc.Cells(sr,4).Value, "#,##0") & " UM"
        scenStr = scenStr & " | Service: " & wsSc.Cells(sr,5).Value & "%" & vbCrLf
    Next sr
    scenStr = scenStr & vbCrLf & "Entrez l'ID du scenario (ex: SC-01):"
    
    Dim chosen As String: chosen = InputBox(scenStr, TOOL_NAME & " â€” Validation Plan", "SC-01")
    If chosen = "" Then Exit Sub
    
    ' Verifier que le scenario existe
    Dim scenRow As Long: scenRow = 0
    For sr = 2 To wsSc.Cells(wsSc.Rows.Count,1).End(xlUp).Row
        If UCase(CStr(wsSc.Cells(sr, 1).Value)) = UCase(chosen) Then scenRow = sr: Exit For
    Next sr
    If scenRow = 0 Then MsgBox "Scenario '" & chosen & "' non trouve.", vbCritical, TOOL_NAME: Exit Sub
    
    ' Confirmation
    Dim conf As String: conf = "Valider le plan avec " & wsSc.Cells(scenRow, 1).Value & " - " & wsSc.Cells(scenRow, 2).Value & " ?" & vbCrLf & _
        "Cout: " & Format(wsSc.Cells(scenRow,4).Value,"#,##0") & " UM | Service: " & wsSc.Cells(scenRow,5).Value & "%"
    If MsgBox(conf, vbYesNo + vbQuestion, TOOL_NAME) <> vbYes Then Exit Sub
    
    ' Ecrire le plan final
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsPf As Worksheet: Set wsPf = ThisWorkbook.Sheets(SH_PLAN_FINAL)
    ClearSheet SH_PLAN_FINAL
    
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    Dim cycleID As String: cycleID = GetCycleID()
    Dim planID  As String: planID  = "PLAN-" & Format(Now(), "YYYYMMDD-HHMMSS")
    Dim valideur As String: valideur = GetCurrentUser()
    
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count,1).End(xlUp).Row
    Dim rowOut As Long: rowOut = 2
    Dim dr As Long
    For dr = 2 To dpLast
        Dim pSku As String: pSku = CStr(wsDp.Cells(dr, 1).Value)
        Dim pMois As Date:  pMois = wsDp.Cells(dr, 2).Value
        Dim pDem  As Double: If IsNumeric(wsDp.Cells(dr,5).Value) Then pDem = CDbl(wsDp.Cells(dr,5).Value)
        
        ' Plant depuis PRODUCT
        Dim pPlant As String: pPlant = "USN-01"
        Dim loProd As ListObject
        On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
        If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
            Dim pp As Long
            For pp = 1 To loProd.DataBodyRange.Rows.Count
                If CStr(loProd.DataBodyRange(pp, 1).Value) = pSku Then
                    pPlant = CStr(loProd.DataBodyRange(pp, 4).Value): Exit For
                End If
            Next pp
        End If
        
        ' Production selon scenario
        Dim prodNorm As Double: prodNorm = pDem
        Dim prodHS   As Double: prodHS   = 0
        Dim prodST   As Double: prodST   = 0
        
        Select Case UCase(chosen)
            Case "SC-02": prodHS = pDem * 0.2
            Case "SC-03": prodST = pDem * 0.1
            Case "SC-04": prodHS = pDem * 0.1: prodST = pDem * 0.05
            Case "SC-05": prodNorm = pDem * 0.9
        End Select
        
        Dim stockFin As Double: stockFin = 0
        Dim nonServi As Double: nonServi = 0
        
        wsPf.Cells(rowOut, 1).Value = planID
        wsPf.Cells(rowOut, 2).Value = 1
        wsPf.Cells(rowOut, 3).Value = pSku
        wsPf.Cells(rowOut, 4).Value = pPlant
        wsPf.Cells(rowOut, 5).Value = pMois
        wsPf.Cells(rowOut, 5).NumberFormat = "mmm-yy"
        wsPf.Cells(rowOut, 6).Value = Round(prodNorm, 0)
        wsPf.Cells(rowOut, 7).Value = Round(prodHS, 0)
        wsPf.Cells(rowOut, 8).Value = Round(prodST, 0)
        wsPf.Cells(rowOut, 9).Value = Round(stockFin, 0)
        wsPf.Cells(rowOut, 10).Value = Round(pDem, 0)
        wsPf.Cells(rowOut, 11).Value = Round(nonServi, 0)
        wsPf.Cells(rowOut, 12).Value = wsSc.Cells(scenRow, 1).Value & " - " & wsSc.Cells(scenRow, 2).Value
        wsPf.Cells(rowOut, 13).Value = "Valide"
        wsPf.Cells(rowOut, 14).Value = valideur
        wsPf.Cells(rowOut, 15).Value = Now()
        wsPf.Cells(rowOut, 15).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsPf.Rows(rowOut).Interior.Color = RGB(200, 240, 200)
        rowOut = rowOut + 1
    Next dr
    
    ' Marquer le scenario comme Valide
    wsSc.Cells(scenRow, 10).Value = "VALIDE"
    wsSc.Rows(scenRow).Interior.Color = RGB(150, 255, 150)
    
    ' Tableau structure
    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsPf.ListObjects("tbl_PLAN_FINAL"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsPf.ListObjects.Add(xlSrcRange, wsPf.Range("A1:O" & (rowOut-1)), , xlYes)
            lo.Name = "tbl_PLAN_FINAL": lo.TableStyle = "TableStyleMedium4"
        End If
    End If
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    SetStatus "Plan VALIDE â€” " & chosen & " â€” " & Now()
    LogMsg "PLANFINAL", "INFO", "Plan valide: " & chosen & " par " & valideur
    
    MsgBox "Plan valide avec " & chosen & "." & vbCrLf & _
           "Consultez PLAN_FINAL." & vbCrLf & _
           "Etape suivante: Exporter vers Power BI.", vbInformation, TOOL_NAME
    wsPf.Activate
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    LogMsg "PLANFINAL", "ERROR", Err.Description
    MsgBox "Erreur validation plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' MONTE CARLO (F18 / EF-23 â€” Expert)
'-------------------------------------------------------------
Public Sub LancerMonteCarlo()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "MC", "INFO", "=== MONTE CARLO ==="
    ShowProgress "Monte Carlo en cours...", 5
    
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    If wsDp.Cells(wsDp.Rows.Count,1).End(xlUp).Row < 2 Then
        MsgBox "Lancez d'abord le Demand Plan.", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    Dim wsMC As Worksheet
    On Error Resume Next: Set wsMC = ThisWorkbook.Sheets("MC_RESULTS"): On Error GoTo 0
    If wsMC Is Nothing Then
        Set wsMC = ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(SH_LOG))
        wsMC.Name = "MC_RESULTS"
        wsMC.Tab.Color = RGB(112, 48, 160)
    End If
    
    ' En-tetes
    Dim mcHdrs() As Variant
    mcHdrs = Array("SKU_ID","Iterations","Moyenne","P50","P90","P_rupture_pct","Cout_moy","Cout_P90","Service_moy_pct","Cycle_ID")
    Dim c As Integer
    For c = 0 To 9: wsMC.Cells(1, c+1).Value = mcHdrs(c): Next c
    wsMC.Range("A1:J1").Font.Bold = $true
    wsMC.Range("A1:J1").Interior.Color = RGB(112, 48, 160)
    wsMC.Range("A1:J1").Font.Color = RGB(255, 255, 255)
    
    Dim nIter As Integer: nIter = CInt(GetParam("MC_ITERATIONS"))
    Dim sigma As Double:  sigma  = CDbl(GetParam("MC_SIGMA_PCT"))
    Dim cycleID As String: cycleID = GetCycleID()
    Dim coutPen As Double: coutPen = 800  ' Cout penalite rupture
    
    ' Lire le Demand Plan par SKU
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count,1).End(xlUp).Row
    Dim skuDem As Object: Set skuDem = CreateObject("Scripting.Dictionary")
    Dim dr As Long
    For dr = 2 To dpLast
        Dim dSku As String: dSku = CStr(wsDp.Cells(dr, 1).Value)
        Dim dQty As Double: If IsNumeric(wsDp.Cells(dr,5).Value) Then dQty = CDbl(wsDp.Cells(dr,5).Value)
        If skuDem.Exists(dSku) Then skuDem(dSku) = skuDem(dSku) + dQty Else skuDem.Add dSku, dQty
    Next dr
    
    ' Limiter a 4 SKU max (Expert)
    Dim skuArr() As String: skuArr = skuDem.Keys()
    Dim nSKU As Integer: nSKU = Application.Min(4, skuDem.Count)
    
    Dim rowOut As Long: rowOut = 2
    Dim s As Integer
    For s = 0 To nSKU - 1
        Dim sk As String: sk = skuArr(s)
        Dim baseDem As Double: baseDem = CDbl(skuDem(sk))
        ShowProgress "Monte Carlo: " & sk, 10 + Int(80 * s / nSKU)
        
        ' Simulation
        Dim simVals() As Double: ReDim simVals(nIter - 1)
        Dim simCouts() As Double: ReDim simCouts(nIter - 1)
        Dim nRupture As Long: nRupture = 0
        Dim totSvc As Double
        
        Dim it As Integer
        For it = 0 To nIter - 1
            ' Box-Muller
            Dim u1 As Double: u1 = Rnd()
            Dim u2 As Double: u2 = Rnd()
            If u1 < 0.0001 Then u1 = 0.0001
            Dim noise As Double: noise = Sqr(-2 * Log(u1)) * Cos(2 * 3.14159265358979 * u2)
            Dim simDem As Double: simDem = Application.Max(0, baseDem * (1 + noise * sigma))
            simVals(it) = simDem
            
            ' Cout si on produit baseDem (pas de levier)
            Dim diff As Double: diff = simDem - baseDem
            Dim cout As Double
            If diff > 0 Then
                cout = diff * coutPen
                nRupture = nRupture + 1
            Else
                cout = Abs(diff) * 12  ' Stockage
            End If
            simCouts(it) = cout
            totSvc = totSvc + Application.Min(1, baseDem / Application.Max(1, simDem))
        Next it
        
        ' Trier pour percentiles
        Dim i As Integer, j As Integer, tmp As Double
        For i = 0 To nIter - 2
            For j = i + 1 To nIter - 1
                If simVals(i) > simVals(j) Then tmp=simVals(i):simVals(i)=simVals(j):simVals(j)=tmp: _
                    tmp=simCouts(i):simCouts(i)=simCouts(j):simCouts(j)=tmp
            Next j
            If i Mod 100 = 0 Then ShowProgress "Monte Carlo tri: " & sk, 50 + Int(30*i/nIter)
        Next i
        
        Dim p50  As Double: p50  = simVals(Int(nIter * 0.5))
        Dim p90  As Double: p90  = simVals(Int(nIter * 0.9))
        Dim cMoy As Double: cMoy = Application.Average(Application.Index(simCouts, 0))
        Dim cP90 As Double: cP90 = simCouts(Int(nIter * 0.9))
        Dim pRup As Double: pRup = nRupture / nIter * 100
        Dim svcMoy As Double: svcMoy = totSvc / nIter * 100
        
        wsMC.Cells(rowOut, 1).Value = sk
        wsMC.Cells(rowOut, 2).Value = nIter
        wsMC.Cells(rowOut, 3).Value = Round(baseDem, 0)
        wsMC.Cells(rowOut, 4).Value = Round(p50, 0)
        wsMC.Cells(rowOut, 5).Value = Round(p90, 0)
        wsMC.Cells(rowOut, 6).Value = Round(pRup, 1)
        wsMC.Cells(rowOut, 7).Value = Round(cMoy, 0)
        wsMC.Cells(rowOut, 8).Value = Round(cP90, 0)
        wsMC.Cells(rowOut, 9).Value = Round(svcMoy, 1)
        wsMC.Cells(rowOut, 10).Value = cycleID
        rowOut = rowOut + 1
    Next s
    
    ' Tableau structure
    Dim lo As ListObject
    On Error Resume Next: Set lo = wsMC.ListObjects("tbl_MC"): On Error GoTo 0
    If lo Is Nothing And rowOut > 2 Then
        Set lo = wsMC.ListObjects.Add(xlSrcRange, wsMC.Range("A1:J" & (rowOut-1)), , xlYes)
        lo.Name = "tbl_MC": lo.TableStyle = "TableStyleMedium8"
    End If
    
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Monte Carlo termine â€” " & Now()
    LogMsg "MC", "INFO", "Monte Carlo OK â€” " & nSKU & " SKU"
    MsgBox "Monte Carlo termine (" & nIter & " iterations)." & vbCrLf & "Consultez MC_RESULTS.", vbInformation, TOOL_NAME
    wsMC.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "MC", "ERROR", Err.Description
    MsgBox "Erreur Monte Carlo: " & Err.Description, vbCritical, TOOL_NAME
End Sub
'@

Set-VBAModule -vbaProject $vbaProject -name "mod06_KPI" -code $vba_KPI

# ============================================================
# MODULE 07 : EXPORT POWER BI + CYCLE COMPLET
# ============================================================
$vba_EXPORT = @'
Option Explicit
'==============================================================
' MODULE EXPORT - Export CSV Power BI + Cycle complet
' F19 / EF-21 + EF-12 â€” AtlasFood S&OP DSS
'==============================================================

'-------------------------------------------------------------
' LANCER LE CYCLE COMPLET (J1 -> J5 automatise)
'-------------------------------------------------------------
Public Sub LancerCycleComplet()
    If MsgBox("Lancer le cycle S&OP complet ?" & vbCrLf & _
              "(Import -> Forecast -> Demand Plan -> Capacity -> Inventory -> Gap -> Scenarios -> KPI)" & vbCrLf & vbCrLf & _
              "Assurez-vous d'avoir importe les donnees (J1) avant.", _
              vbYesNo + vbQuestion, TOOL_NAME) <> vbYes Then Exit Sub
    
    LogMsg "CYCLE", "INFO", "=== DEBUT CYCLE COMPLET ==="
    ShowProgress "Cycle complet: Forecast...", 5
    Call LancerForecast
    
    ShowProgress "Cycle complet: Demand Plan...", 25
    Call LancerDemandPlan
    
    ShowProgress "Cycle complet: Capacity Plan...", 45
    Call LancerCapacityPlan
    
    ShowProgress "Cycle complet: Inventory Plan...", 60
    Call LancerInventoryPlan
    
    ShowProgress "Cycle complet: Gap Analysis...", 75
    Call LancerGapAnalysis
    
    ShowProgress "Cycle complet: Scenarios...", 85
    Call GenererScenarios
    
    ShowProgress "Cycle complet: KPI...", 95
    Call CalculerKPI
    
    ClearProgress
    LogMsg "CYCLE", "INFO", "=== CYCLE COMPLET TERMINE ==="
    SetStatus "Cycle complet termine â€” " & Now()
    MsgBox "Cycle S&OP complet execute avec succes!" & vbCrLf & _
           "Toutes les etapes J2 a J4 ont ete calculees." & vbCrLf & vbCrLf & _
           "Consultez maintenant les scenarios et validez le plan (J5).", vbInformation, TOOL_NAME
    ThisWorkbook.Sheets(SH_SCENARIOS).Activate
End Sub

'-------------------------------------------------------------
' EXPORT CSV VERS POWER BI (F19 / EF-21)
'-------------------------------------------------------------
Public Sub ExporterVersPowerBI()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    LogMsg "EXPORT", "INFO", "=== EXPORT POWER BI ==="
    ShowProgress "Export Power BI...", 5
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Dossier de destination pour les CSV Power BI"
    fd.InitialFileName = ThisWorkbook.Path & "\..\PowerBI\"
    
    If fd.Show <> -1 Then GoTo CleanExit
    Dim exportDir As String: exportDir = fd.SelectedItems(1)
    
    Dim wsExp As Worksheet: Set wsExp = ThisWorkbook.Sheets(SH_EXPORT)
    Dim expRow As Long: expRow = 2
    
    ' Effacer le tableau de bord des exports
    Dim expLast As Long: expLast = wsExp.Cells(wsExp.Rows.Count,1).End(xlUp).Row
    If expLast > 1 Then wsExp.Range("A2:D" & expLast).ClearContents
    
    ' Tables a exporter
    Dim tablesToExport() As Variant
    tablesToExport = Array( _
        Array(SH_FORECAST,     "tbl_FORECAST",   "PBI_Forecast.csv"), _
        Array(SH_DEMAND_PLAN,  "tbl_DEMAND_PLAN","PBI_DemandPlan.csv"), _
        Array(SH_CAPACITY,     "tbl_CAPACITY",   "PBI_CapacityPlan.csv"), _
        Array(SH_INVENTORY,    "tbl_INVENTORY",  "PBI_InventoryPlan.csv"), _
        Array(SH_GAP,          "tbl_GAP",        "PBI_GapAnalysis.csv"), _
        Array(SH_SCENARIOS,    "tbl_SCENARIOS",  "PBI_Scenarios.csv"), _
        Array(SH_PLAN_FINAL,   "tbl_PLAN_FINAL", "PBI_PlanFinal.csv"), _
        Array(SH_KPI,          "tbl_KPI",        "PBI_KPI.csv"), _
        Array(SH_ALERTS,       "",               "PBI_Alertes.csv"), _
        Array(SH_PRODUCT,      "tbl_PRODUCT",    "PBI_Product.csv"), _
        Array(SH_PLANT,        "tbl_PLANT",      "PBI_Plant.csv") _
    )
    
    Dim t As Integer
    For t = 0 To UBound(tablesToExport)
        ShowProgress "Export: " & tablesToExport(t)(2), 5 + Int(90 * t / (UBound(tablesToExport)+1))
        
        Dim shName  As String: shName  = CStr(tablesToExport(t)(0))
        Dim csvName As String: csvName = CStr(tablesToExport(t)(2))
        Dim csvPath As String: csvPath = exportDir & "\" & csvName
        
        Dim ws As Worksheet
        On Error Resume Next: Set ws = ThisWorkbook.Sheets(shName): On Error GoTo 0
        If ws Is Nothing Then
            wsExp.Cells(expRow, 1).Value = csvName
            wsExp.Cells(expRow, 4).Value = "ERREUR: feuille manquante"
            expRow = expRow + 1
            GoTo NextTable
        End If
        
        ' Ecrire CSV UTF-8
        Dim nbRows As Long
        nbRows = ExportFeuilleCSV(ws, csvPath)
        
        wsExp.Cells(expRow, 1).Value = csvName
        wsExp.Cells(expRow, 2).Value = Now()
        wsExp.Cells(expRow, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsExp.Cells(expRow, 3).Value = nbRows
        wsExp.Cells(expRow, 4).Value = IIf(nbRows > 0, "OK", "Vide")
        If nbRows > 0 Then wsExp.Rows(expRow).Interior.Color = RGB(200,240,200) _
        Else wsExp.Rows(expRow).Interior.Color = RGB(255,220,150)
        expRow = expRow + 1
        LogMsg "EXPORT", "INFO", csvName & " -> " & nbRows & " lignes"
NextTable:
    Next t
    
    Application.ScreenUpdating = True
    ClearProgress: SetStatus "Export Power BI termine â€” " & Now()
    
    MsgBox "Export termine!" & vbCrLf & (UBound(tablesToExport)+1) & " fichiers CSV crees dans:" & vbCrLf & exportDir & vbCrLf & vbCrLf & _
           "Actualisez votre rapport Power BI pour voir les donnees mises a jour.", vbInformation, TOOL_NAME
    wsExp.Activate
    Exit Sub
    
CleanExit:
    Application.ScreenUpdating = True: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: ClearProgress
    LogMsg "EXPORT", "ERROR", Err.Description
    MsgBox "Erreur Export: " & Err.Description, vbCritical, TOOL_NAME
End Sub

' Exporter une feuille en CSV UTF-8 via ADODB.Stream
Private Function ExportFeuilleCSV(ws As Worksheet, csvPath As String) As Long
    ExportFeuilleCSV = 0
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 1 Or lastCol < 1 Then Exit Function
    
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Open
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    
    Dim r As Long, c As Long
    For r = 1 To lastRow
        Dim lineStr As String: lineStr = ""
        For c = 1 To lastCol
            Dim cellV As String
            If ws.Cells(r, c).NumberFormat Like "*m/d*" Or ws.Cells(r, c).NumberFormat Like "*yyyy*" Then
                If IsDate(ws.Cells(r, c).Value) Then
                    cellV = Format(ws.Cells(r, c).Value, "yyyy-mm-dd")
                Else
                    cellV = CStr(ws.Cells(r, c).Value)
                End If
            Else
                cellV = CStr(ws.Cells(r, c).Value)
            End If
            ' Echapper les guillemets
            cellV = Replace(cellV, """", """""")
            If InStr(cellV, ";") > 0 Or InStr(cellV, """") > 0 Then cellV = """" & cellV & """"
            If c > 1 Then lineStr = lineStr & ";"
            lineStr = lineStr & cellV
        Next c
        stm.WriteText lineStr & vbCrLf
    Next r
    
    ' Supprimer le BOM UTF-8 et sauvegarder
    stm.Position = 0
    stm.Type = 1  ' adTypeBinary
    stm.Position = 3  ' Sauter le BOM
    
    Dim stm2 As Object: Set stm2 = CreateObject("ADODB.Stream")
    stm2.Open
    stm2.Type = 1
    stm2.Write stm.Read()
    stm2.SaveToFile csvPath, 2  ' adSaveCreateOverWrite
    stm2.Close
    stm.Close
    
    ExportFeuilleCSV = lastRow - 1  ' Exclure l'entete
End Function

'-------------------------------------------------------------
' EFFACER TOUS LES CALCULS (recommencer)
'-------------------------------------------------------------
Public Sub ReinitialiserCalculs()
    If MsgBox("Effacer tous les calculs (Forecast, Plans, Gap, Scenarios, KPI, Plan Final) ?" & vbCrLf & _
              "Les donnees importees seront conservees.", _
              vbYesNo + vbExclamation, TOOL_NAME) <> vbYes Then Exit Sub
    
    Dim sheetsToClean() As Variant
    sheetsToClean = Array(SH_FORECAST, SH_DEMAND_PLAN, SH_CAPACITY, SH_INVENTORY, SH_GAP, SH_SCENARIOS, SH_PLAN_FINAL, SH_KPI, SH_ALERTS)
    
    Dim s As Variant
    For Each s In sheetsToClean
        ClearSheet CStr(s)
    Next s
    
    ' Effacer MC_RESULTS si existe
    On Error Resume Next
    Dim wsMC As Worksheet: Set wsMC = ThisWorkbook.Sheets("MC_RESULTS")
    If Not wsMC Is Nothing Then ClearSheet "MC_RESULTS"
    On Error GoTo 0
    
    SetStatus "Calculs reinitialises â€” " & Now()
    LogMsg "SYSTEM", "INFO", "Reinitialisation des calculs effectuee."
    MsgBox "Calculs reinitialises. Relancez le cycle depuis J2 (Forecast).", vbInformation, TOOL_NAME
End Sub

'-------------------------------------------------------------
' ALLER A UNE FEUILLE
'-------------------------------------------------------------
Public Sub NaviguerVers(shName As String)
    On Error Resume Next
    ThisWorkbook.Sheets(shName).Activate
    On Error GoTo 0
End Sub
'@

Set-VBAModule -vbaProject $vbaProject -name "mod07_EXPORT" -code $vba_EXPORT

# ============================================================
# CODE DE LA FEUILLE ACCUEIL (Sheet Events)
# ============================================================
$vba_ACCUEIL = @'
Option Explicit

' Bouton: Importer les donnees (J1)
Private Sub btn_Importer_les_donn_es__J1__Click()
    Call ImporterDonnees
End Sub

' Bouton: Lancer Forecast (J2)
Private Sub btn_Lancer_Forecast__J2__Click()
    Call LancerForecast
End Sub

' Bouton: Demand Plan (J2b)
Private Sub btn_Demand_Plan__J2b__Click()
    Call LancerDemandPlan
End Sub

' Bouton: Capacity + Inventory + Gap (J3)
Private Sub btn_Capacity___Inventory___Gap__J3__Click()
    Call LancerCapacityPlan
    Call LancerInventoryPlan
    Call LancerGapAnalysis
End Sub

' Bouton: Generer Scenarios (J4)
Private Sub btn_G_n_rer_Sc_narios__J4__Click()
    Call GenererScenarios
End Sub

' Bouton: Calculer KPI (J4b)
Private Sub btn_Calculer_KPI__J4b__Click()
    Call CalculerKPI
End Sub

' Bouton: Valider Plan Final (J5)
Private Sub btn_Valider_Plan_Final__J5__Click()
    Call ValiderPlanFinal
End Sub

' Bouton: Monte Carlo (Expert)
Private Sub btn_Monte_Carlo__Expert__Click()
    Call LancerMonteCarlo
End Sub

' Bouton: Exporter vers Power BI
Private Sub btn_Exporter_vers_Power_BI_Click()
    Call ExporterVersPowerBI
End Sub

' Bouton: Cycle complet
Private Sub btn_Cycle_Complet_Click()
    Call LancerCycleComplet
End Sub

' Bouton: Reinitialiser
Private Sub btn_R_initialiser_Click()
    Call ReinitialiserCalculs
End Sub

' Bouton: Voir LOG
Private Sub btn_Voir_LOG_Click()
    ThisWorkbook.Sheets(SH_LOG).Activate
End Sub
'@

# Injecter dans la feuille ACCUEIL
$wsAccSheet = $wb.VBProject.VBComponents["ACCUEIL"]
if ($null -eq $wsAccSheet) {
    # Chercher par type (feuille)
    foreach ($vbc in $wb.VBProject.VBComponents) {
        if ($vbc.Properties("Name").Value -eq "ACCUEIL" -or $vbc.Name -eq "Sheet1") {
            $wsAccSheet = $vbc; break
        }
    }
}
if ($wsAccSheet) {
    $wsAccSheet.CodeModule.DeleteLines(1, $wsAccSheet.CodeModule.CountOfLines)
    $wsAccSheet.CodeModule.InsertLines(1, $vba_ACCUEIL)
}

# ============================================================
# CODE ThisWorkbook (ouverture / fermeture)
# ============================================================
$vba_THISWORKBOOK = @'
Option Explicit

Private Sub Workbook_Open()
    ' Aller a ACCUEIL
    ThisWorkbook.Sheets(SH_ACCUEIL).Activate
    
    ' Message de bienvenue
    Dim msg As String
    msg = "Bienvenue dans " & TOOL_NAME & " v" & TOOL_VERSION & vbCrLf & vbCrLf & _
          "Pour commencer:" & vbCrLf & _
          "1. Saisissez votre nom dans la cellule Utilisateur (colonne C, ligne 9)" & vbCrLf & _
          "2. Cliquez sur 'Importer les donnees' (J1)" & vbCrLf & _
          "3. Suivez le workflow S&OP de gauche a droite" & vbCrLf & vbCrLf & _
          "Source des donnees: dossier /Data (CSV)" & vbCrLf & _
          "Documentation: dossier /Documentation"
    
    MsgBox msg, vbInformation, TOOL_NAME
    
    LogMsg "SYSTEM", "INFO", "Ouverture du classeur par: " & GetCurrentUser()
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogMsg "SYSTEM", "INFO", "Sauvegarde par: " & GetCurrentUser()
End Sub
'@

$tbComp = $wb.VBProject.VBComponents["ThisWorkbook"]
if ($tbComp) {
    $tbComp.CodeModule.DeleteLines(1, $tbComp.CodeModule.CountOfLines)
    $tbComp.CodeModule.InsertLines(1, $vba_THISWORKBOOK)
}

Write-Host "   -> Modules VBA injectes."

# ================================================================
# ETAPE 7 : AJOUTER LES BOUTONS ACTIVEX SUR LA FEUILLE ACCUEIL
# ================================================================
Write-Host "[7/10] Ajout des boutons interface..."

$wsAcc = $wb.Sheets["ACCUEIL"]

# Configuration des boutons: caption, top, left, width, height, color
$buttons = @(
    @{c="Importer les donnees (J1)";    t=240; l=30;  w=210; h=40; bg=0x2E75B6},
    @{c="Lancer Forecast (J2)";         t=240; l=260; w=190; h=40; bg=0x70AD47},
    @{c="Demand Plan (J2b)";            t=240; l=470; w=180; h=40; bg=0x70AD47},
    @{c="Capacity + Inventory + Gap (J3)"; t=240; l=670; w=230; h=40; bg=0xFFC000},
    @{c="Generer Scenarios (J4)";       t=300; l=30;  w=190; h=40; bg=0x7030A0},
    @{c="Calculer KPI (J4b)";           t=300; l=240; w=180; h=40; bg=0x7030A0},
    @{c="Valider Plan Final (J5)";      t=300; l=440; w=200; h=40; bg=0x00B050},
    @{c="Monte Carlo (Expert)";         t=300; l=660; w=190; h=40; bg=0xC00000},
    @{c="Cycle Complet";                t=370; l=30;  w=160; h=40; bg=0xED7D31},
    @{c="Exporter vers Power BI";       t=370; l=210; w=200; h=40; bg=0x595959},
    @{c="Reinitialiser";                t=370; l=430; w=150; h=40; bg=0xC00000},
    @{c="Voir LOG";                     t=370; l=600; w=130; h=40; bg=0x595959}
)

foreach ($b in $buttons) {
    try {
        $oleBtn = $wsAcc.OLEObjects.Add(
            "Forms.CommandButton.1",
            [System.Reflection.Missing]::Value,
            $false, $false,
            [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value,
            $b.l, $b.t, $b.w, $b.h
        )
        $oleBtn.Object.Caption   = $b.c
        $oleBtn.Object.BackColor = $b.bg
        $oleBtn.Object.ForeColor = 0xFFFFFF
        $oleBtn.Object.Font.Bold = $true
        $oleBtn.Object.Font.Size = 9
        # Nom du bouton (pour liaison avec macro)
        $safeCaption = $b.c -replace "[^A-Za-z0-9]","_"
        $oleBtn.Name = "btn_" + $safeCaption
    } catch {
        Write-Host "   [WARN] Bouton '" + $b.c + "' non ajoute: " + $_.Exception.Message
    }
}

# Masquer les feuilles techniques pour l'utilisateur final
$hiddenSheets = @("LOG", "EXPORT_POWERBI")
foreach ($shHide in $hiddenSheets) {
    try { $wb.Sheets[$shHide].Visible = 2 } catch {}  # xlSheetVeryHidden
}

Write-Host "   -> Boutons ajoutes."

# ================================================================
# ETAPE 8 : MISE EN PAGE FINALE ET PROTECTION LEGERE
# ================================================================
Write-Host "[8/10] Mise en page finale..."

# Colonne A (index) masquee sur ACCUEIL
$wsAcc.Columns("A").Hidden = $true
$wsAcc.Columns("O:Z").Hidden = $true

# Figer la ligne 1 sur les feuilles de donnees
$dataSheetsList = @("SALES_HISTORY","PRODUCTION_HISTORY","DELIVERY_HISTORY","FORECAST","DEMAND_PLAN","CAPACITY_PLAN","INVENTORY_PLAN","GAP_ANALYSIS","SCENARIOS","PLAN_FINAL","KPI_RESULTS","ALERTES")
foreach ($dsn in $dataSheetsList) {
    try {
        $wsd = $wb.Sheets[$dsn]
        $wsd.Application.ActiveWindow.SplitRow = 1
        $wsd.Activate()
        $wb.Application.ActiveWindow.FreezePanes = $true
    } catch {}
}

# Revenir sur ACCUEIL
$wsAcc.Activate()

Write-Host "   -> Mise en page OK."

# ================================================================
# ETAPE 9 : SAUVEGARDER EN XLSM
# ================================================================
Write-Host "[9/10] Sauvegarde du fichier .xlsm..."

try {
    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($xlsxPath, 52)
    Write-Host "   -> Sauvegarde OK: $xlsxPath"
} catch {
    Write-Host "ERREUR sauvegarde: $($_.Exception.Message)"
    # Essayer de sauvegarder quand meme
    $xlsxPath2 = $xlsxPath -replace ".xlsm",".xlsx"
    try { $wb.SaveAs($xlsxPath2, 51) } catch {}
}

# ================================================================
# ETAPE 10 : FERMER EXCEL
# ================================================================
Write-Host "[10/10] Fermeture d'Excel..."
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "========================================================"
Write-Host "  CONSTRUCTION TERMINEE"
Write-Host "========================================================"
Write-Host "  Fichier: $xlsxPath"
Write-Host ""
Write-Host "  Pour utiliser l'outil:"
Write-Host "  1. Ouvrez AtlasFood_SOP.xlsm"
Write-Host "  2. Activez les macros"
Write-Host "  3. Suivez le workflow sur la feuille ACCUEIL"
Write-Host "========================================================"



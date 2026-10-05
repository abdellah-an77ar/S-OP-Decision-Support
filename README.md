============================================================
README - AtlasFood S&OP Decision Support System v1.0
============================================================

POUR DEMARRER IMMEDIATEMENT
-----------------------------
1. Ouvrir : Application/AtlasFood_SOP.xlsm
2. Activer les macros
3. Cliquer : [Importer les donnees] -> selectionner le dossier /Data/
4. Suivre le workflow J1 -> J2 -> J3 -> J4 -> J5

STRUCTURE DU PROJET
--------------------
/Application/
    AtlasFood_SOP.xlsm        <- FICHIER PRINCIPAL
    build_xlsm.ps1            <- Script de reconstruction
    add_buttons.ps1           <- Script correctif boutons

/Data/
    SALES_HISTORY.csv         <- 795 lignes, 5 SKU, janv-aout 2023
    PRODUCTION_HISTORY.csv    <- 795 lignes
    DELIVERY_HISTORY.csv      <- 795 lignes
    PRODUCT.csv               <- 5 SKU (3 familles)
    PLANT.csv                 <- 3 usines
    CALENDAR.csv              <- 8 mois 2023
    COST.csv                  <- 18 lignes de couts
    generate_data.ps1         <- Script de regeneration des donnees

/Documentation/
    Guide_Utilisateur.txt     <- Manuel complet
    Guide_PowerBI.txt         <- Connexion Power BI Desktop
    Historique_Construction.txt <- Decisions techniques
    README.txt                <- Ce fichier

/PowerBI/                     <- Vide - rempli par "Exporter vers Power BI"
    (PBI_*.csv seront crees ici apres export)

/VBA_Modules/                 <- Sources VBA (versionning)
    mod01_CONFIG.bas
    mod02_IMPORT.bas
    mod03_FORECAST.bas
    mod04_PLANS.bas
    mod05_SCENARIOS.bas
    mod06_KPI.bas
    mod07_EXPORT.bas
    ThisWorkbook.bas

FONCTIONNALITES
----------------
[J1] Import CSV UTF-8 via ADODB.Stream
[J2] Forecast (MM, Tendance, Saisonnalite) + back-test + WAPE
[J2] Demand Plan avec ajustements manuels
[J3] Capacity Plan par usine/mois
[J3] Inventory Plan avec stock de securite statistique
[J3] Gap Analysis (deficit/surplus capacitaire)
[J4] 5 scenarios compares (Baseline, HS, ST, Mixte, Lissage)
[J4] 8 KPI avec alertes automatiques
[J5] Validation et generation du Plan Final
[Expert] Monte Carlo 1000 iterations (robustesse)
[Export] 11 fichiers CSV UTF-8 pour Power BI

TECHNOLOGIES
-------------
Microsoft Excel + VBA (COM Automation)
Microsoft Power BI (reporting - voir Guide_PowerBI.txt)

DONNEES
--------
Donnees synthetiques - Source: inspire de SupplyGraph
(Wasi et al., 2024 - 2401.15299)
Periode : Jan-Aout 2023 | 5 SKU | 3 Usines

============================================================

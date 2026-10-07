VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Parametres 
   Caption         =   "Parametres Globaux & Structure des Couts (T1)"
   ClientHeight    =   8340.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9180.001
   OleObjectBlob   =   "frm_Parametres.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Parametres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' In-memory storage for multi-plant costs and calendar/capacities
Private m_Costs As Object
Private m_CurrentPlantCost As String
Private m_CurrentPlantCap As String
Private m_CurrentMois As String
Private m_CalCap As Object
Private m_Initializing As Boolean

Private Sub UserForm_Initialize()
    m_Initializing = True
    InitialiserCaptions
    InitialiserFormulaire
    m_Initializing = False
End Sub

Private Sub InitialiserCaptions()
    Me.Caption = "Param" & Chr(232) & "tres Globaux & Structure des Co" & Chr(251) & "ts (T1)"
    lbl_Title.Caption = "  GOUVERNANCE, PARAM" & Chr(200) & "TRES & STRUCTURE DES CO" & Chr(219) & "TS S&OP"
    
    ' MultiPage Tabs
    mp_Tabs.Pages.Item(0).Caption = "1. Politique & Cibles S&OP"
    mp_Tabs.Pages.Item(1).Caption = "2. Structure des Co" & Chr(251) & "ts"
    mp_Tabs.Pages.Item(2).Caption = "3. Calendrier & Capacit" & Chr(233) & "s"
    
    ' Tab 1: Politique & Cibles
    lbl_NiveauService.Caption = "Taux de service client cible (%) :"
    lbl_NiveauService_Desc.Caption = "% (cible client, d" & Chr(233) & "faut 95%)"
    lbl_SeuilAttention.Caption = "Seuil d'alerte Attention / Orange (%) :"
    lbl_SeuilAttention_Desc.Caption = "% (charge usine, d" & Chr(233) & "faut 85%)"
    lbl_SeuilSurcharge.Caption = "Seuil d'alerte Surcharge / Rouge (%) :"
    lbl_SeuilSurcharge_Desc.Caption = "% (saturation usine, d" & Chr(233) & "faut 100%)"
    lbl_SeuilDLC.Caption = "Seuil DLC restante critique (Jours) :"
    lbl_SeuilDLC_Desc.Caption = "jours (alerte p" & Chr(233) & "remption, d" & Chr(233) & "faut 30)"
    lbl_StockMax.Caption = "Plafond Stock Max Global :"
    lbl_StockMax_Desc.Caption = "unit" & Chr(233) & "s (capacit" & Chr(233) & " max r" & Chr(233) & "seau)"
    lbl_DelaiLivraison.Caption = "D" & Chr(233) & "lai de r" & Chr(233) & "approvisionnement (Jours) :"
    lbl_DelaiLivraison_Desc.Caption = "jours (d" & Chr(233) & "lai moyen, d" & Chr(233) & "faut 5)"
    lbl_MargeCapacite.Caption = "Marge flexibilit" & Chr(233) & " capacit" & Chr(233) & " (%) :"
    lbl_MargeCapacite_Desc.Caption = "% (flexibilit" & Chr(233) & " r" & Chr(233) & "seau, d" & Chr(233) & "faut 10%)"
    lbl_Tab1Help.Caption = "Gouvernance S&OP (EMINES - Section 5.1 & T1) : Aucune hypoth" & Chr(232) & "se n'est cod" & Chr(233) & "e en dur. Le taux de service d" & Chr(233) & "termine le facteur de s" & Chr(233) & "curit" & Chr(233) & " Z (loi normale). Les seuils d'alerte gouvernent le Gap Analysis."
    
    ' Tab 2: Structure des Couts
    lbl_CostPlant.Caption = "Usine (Plant_ID) :"
    lbl_CostProd.Caption = "Production standard :"
    lbl_UnitProd.Caption = "UM/u"
    lbl_CostHS.Caption = "Heures Sup (HS) :"
    lbl_UnitHS.Caption = "UM/u"
    lbl_CostST.Caption = "Sous-traitance (ST) :"
    lbl_UnitST.Caption = "UM/u"
    lbl_CostStockage.Caption = "Possession / Stockage :"
    lbl_UnitStockage.Caption = "UM/u/m"
    lbl_CostPenurie.Caption = "P" & Chr(233) & "nurie / Rupture :"
    lbl_UnitPenurie.Caption = "UM/u"
    lbl_CostPeremption.Caption = "D" & Chr(233) & "pr" & Chr(233) & "ciation DLC :"
    lbl_UnitPeremption.Caption = "UM/u"
    lbl_CostSummary.Caption = "Matrice des Co" & Chr(251) & "ts : Co" & Chr(251) & "ts unitaires appliqu" & Chr(233) & "s lors de la valorisation des sc" & Chr(233) & "narios S&OP (Standard, HS, ST, Stockage, Ruptures et Pertes DLC). S" & Chr(233) & "lectionnez chaque usine pour personnaliser ses bar" & Chr(232) & "mes."
    
    ' Tab 3: Calendrier & Capacites
    lbl_MoisCal.Caption = "Mois d'horizon :"
    lbl_PlantCap.Caption = "Usine concern" & Chr(233) & "e :"
    lbl_JoursOuvres.Caption = "Jours ouvr" & Chr(233) & "s / mois :"
    lbl_JoursArret.Caption = "Arr" & Chr(234) & "ts maintenance :"
    lbl_CapNominale.Caption = "Capacit" & Chr(233) & " nominale :"
    lbl_MargeCapPlant.Caption = "Marge capacit" & Chr(233) & " (%) :"
    lbl_CapMaxHS.Caption = "Plafond max HS (%) :"
    lbl_CapMaxST.Caption = "Plafond max ST (u) :"
    lbl_BoxCapTitle.Caption = "Estimation de la Capacit" & Chr(233) & " Nette Disponible :"
    lbl_CapFormule.Caption = "Formule de calcul : Cap. Nette = Nominale * (1 + Marge) * (Jours Ouvr" & Chr(233) & "s - Arr" & Chr(234) & "ts) / Jours Ouvr" & Chr(233) & "s + Max HS + Max ST."
    
    ' Bottom Buttons
    lbl_Info.Caption = "Pr" & Chr(234) & "t pour modification des param" & Chr(232) & "tres."
    btn_ChargerDefaut.Caption = "R" & Chr(233) & "tablir D" & Chr(233) & "fauts"
    btn_Enregistrer.Caption = "Enregistrer Param" & Chr(232) & "tres"
    btn_Fermer.Caption = "Fermer"
End Sub

Private Sub InitialiserFormulaire()
    Dim wsCal As Worksheet
    Dim tblCal As Object
    Dim rCal As Long
    Dim mVal As String
    Dim sSurch As Variant, sDlc As Variant, sStkMax As Variant, sDel As Variant, sMar As Variant
    
    Set m_Costs = CreateObject("Scripting.Dictionary")
    Set m_CalCap = CreateObject("Scripting.Dictionary")
    
    ' 1. Charger Tab 1: Parametres Globaux
    txt_NiveauService.Text = CStr(Round(CDbl(GetParam("NIVEAU_SERVICE")) * 100#, 1))
    txt_SeuilAttention.Text = CStr(Round(CDbl(GetParam("SEUIL_ALERTE_GAP")) * 100#, 1))
    
    sSurch = GetParam("SEUIL_SURCHARGE_CAP")
    If IsNumeric(sSurch) Then txt_SeuilSurcharge.Text = CStr(Round(CDbl(sSurch) * 100#, 1)) Else txt_SeuilSurcharge.Text = "100"
    
    sDlc = GetParam("SEUIL_EXPIRATION_DLC")
    If IsNumeric(sDlc) Then txt_SeuilDLC.Text = CStr(sDlc) Else txt_SeuilDLC.Text = "30"
    
    sStkMax = GetParam("STOCK_MAX_GLOBAL")
    If IsNumeric(sStkMax) Then txt_StockMax.Text = CStr(sStkMax) Else txt_StockMax.Text = "50000"
    
    sDel = GetParam("DELAI_LIVRAISON")
    If IsNumeric(sDel) Then txt_DelaiLivraison.Text = CStr(sDel) Else txt_DelaiLivraison.Text = "5"
    
    sMar = GetParam("MARGE_CAPACITE")
    If IsNumeric(sMar) Then txt_MargeCapacite.Text = CStr(Round(CDbl(sMar) * 100#, 1)) Else txt_MargeCapacite.Text = "10"
    
    ' 2. Charger Tab 2: Matrice des Couts
    cbo_CostPlant.Clear
    cbo_CostPlant.AddItem "USN-01"
    cbo_CostPlant.AddItem "USN-02"
    cbo_CostPlant.AddItem "USN-03"
    
    ChargerTousLesCoutsDepuisTable
    cbo_CostPlant.ListIndex = 0
    m_CurrentPlantCost = "USN-01"
    AfficherCoutsPourUsine "USN-01"
    
    ' 3. Charger Tab 3: Calendrier & Capacites
    cbo_PlantCap.Clear
    cbo_PlantCap.AddItem "USN-01"
    cbo_PlantCap.AddItem "USN-02"
    cbo_PlantCap.AddItem "USN-03"
    
    cbo_MoisCal.Clear
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets(SH_CALENDAR)
    If Not wsCal Is Nothing Then
        Set tblCal = wsCal.ListObjects("tbl_CALENDAR")
        If Not tblCal Is Nothing Then
            If Not tblCal.DataBodyRange Is Nothing Then
                For rCal = 1 To tblCal.DataBodyRange.Rows.Count
                    mVal = Format(tblCal.DataBodyRange(rCal, 2).Value, "YYYY-MM")
                    If Len(mVal) = 7 Then cbo_MoisCal.AddItem mVal
                Next rCal
            End If
        End If
    End If
    On Error GoTo 0
    
    If cbo_MoisCal.ListCount = 0 Then
        cbo_MoisCal.AddItem "2023-06"
        cbo_MoisCal.AddItem "2023-07"
        cbo_MoisCal.AddItem "2023-08"
        cbo_MoisCal.AddItem "2023-09"
        cbo_MoisCal.AddItem "2023-10"
        cbo_MoisCal.AddItem "2023-11"
    End If
    
    ChargerToutesCapacitesDepuisTables
    cbo_PlantCap.ListIndex = 0
    cbo_MoisCal.ListIndex = 0
    m_CurrentPlantCap = cbo_PlantCap.Text
    m_CurrentMois = cbo_MoisCal.Text
    AfficherCapacitePourSelection m_CurrentPlantCap, m_CurrentMois
End Sub

' -------------------------------------------------------------
' GESTION TAB 2: MATRICE DES COUTS
' -------------------------------------------------------------
Private Sub ChargerTousLesCoutsDepuisTable()
    Dim wsC As Worksheet
    Dim tblCost As Object
    Dim r As Long
    Dim cType As String, cPlant As String, k As String
    Dim cMontant As Double
    
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets(SH_COST)
    If wsC Is Nothing Then Exit Sub
    Set tblCost = wsC.ListObjects("tbl_COST")
    If tblCost Is Nothing Then Exit Sub
    If tblCost.DataBodyRange Is Nothing Then Exit Sub
    
    For r = 1 To tblCost.DataBodyRange.Rows.Count
        cType = Trim(CStr(tblCost.DataBodyRange(r, 1).Value))
        cPlant = Trim(CStr(tblCost.DataBodyRange(r, 2).Value))
        cMontant = Val(Replace(CStr(tblCost.DataBodyRange(r, 3).Value), ",", "."))
        k = cPlant & "|" & cType
        m_Costs(k) = cMontant
    Next r
    On Error GoTo 0
End Sub

Private Sub SauvegarderCoutsChampsEnMemoire(ByVal plantId As String)
    If plantId = "" Then Exit Sub
    If IsNumeric(txt_CostProd.Text) Then m_Costs(plantId & "|Prod") = Val(Replace(txt_CostProd.Text, ",", "."))
    If IsNumeric(txt_CostHS.Text) Then m_Costs(plantId & "|HS") = Val(Replace(txt_CostHS.Text, ",", "."))
    If IsNumeric(txt_CostST.Text) Then m_Costs(plantId & "|ST") = Val(Replace(txt_CostST.Text, ",", "."))
    If IsNumeric(txt_CostStockage.Text) Then m_Costs(plantId & "|Stockage") = Val(Replace(txt_CostStockage.Text, ",", "."))
    If IsNumeric(txt_CostPenurie.Text) Then m_Costs(plantId & "|Penurie") = Val(Replace(txt_CostPenurie.Text, ",", "."))
    If IsNumeric(txt_CostPeremption.Text) Then m_Costs(plantId & "|Peremption") = Val(Replace(txt_CostPeremption.Text, ",", "."))
End Sub

Private Sub AfficherCoutsPourUsine(ByVal plantId As String)
    Select Case plantId
        Case "USN-01": lbl_PlantNom.Caption = "Usine Nord (Casablanca)"
        Case "USN-02": lbl_PlantNom.Caption = "Usine Sud (Agadir)"
        Case "USN-03": lbl_PlantNom.Caption = "Usine Est (Oujda)"
        Case Else:     lbl_PlantNom.Caption = plantId
    End Select
    
    txt_CostProd.Text = ObtenirCoutVal(plantId, "Prod", 150)
    txt_CostHS.Text = ObtenirCoutVal(plantId, "HS", 350)
    txt_CostST.Text = ObtenirCoutVal(plantId, "ST", 500)
    txt_CostStockage.Text = ObtenirCoutVal(plantId, "Stockage", 12)
    txt_CostPenurie.Text = ObtenirCoutVal(plantId, "Penurie", 800)
    txt_CostPeremption.Text = ObtenirCoutVal(plantId, "Peremption", 200)
End Sub

Private Function ObtenirCoutVal(ByVal plantId As String, ByVal cType As String, ByVal defVal As Double) As String
    Dim k As String: k = plantId & "|" & cType
    If m_Costs.Exists(k) Then
        ObtenirCoutVal = CStr(m_Costs(k))
    Else
        ObtenirCoutVal = CStr(defVal)
    End If
End Function

Private Sub cbo_CostPlant_Change()
    If m_Initializing Then Exit Sub
    If m_CurrentPlantCost <> "" Then
        SauvegarderCoutsChampsEnMemoire m_CurrentPlantCost
    End If
    m_CurrentPlantCost = cbo_CostPlant.Text
    AfficherCoutsPourUsine m_CurrentPlantCost
End Sub

' -------------------------------------------------------------
' GESTION TAB 3: CALENDRIER & CAPACITES
' -------------------------------------------------------------
Private Sub ChargerToutesCapacitesDepuisTables()
    Dim wsCal As Worksheet
    Dim tblCal As Object
    Dim wsPl As Worksheet
    Dim tblPl As Object
    Dim cr As Long, pr As Long
    Dim mDate As String, pid As String
    Dim jo As Double, ja As Double
    Dim cRef As Double, mg As Double, hs As Double, st As Double
    
    On Error Resume Next
    
    ' Charger CALENDAR par mois
    Set wsCal = ThisWorkbook.Sheets(SH_CALENDAR)
    If Not wsCal Is Nothing Then
        Set tblCal = wsCal.ListObjects("tbl_CALENDAR")
        If Not tblCal Is Nothing Then
            If Not tblCal.DataBodyRange Is Nothing Then
                For cr = 1 To tblCal.DataBodyRange.Rows.Count
                    mDate = Format(tblCal.DataBodyRange(cr, 2).Value, "YYYY-MM")
                    jo = Val(Replace(CStr(tblCal.DataBodyRange(cr, 1).Value), ",", "."))
                    ja = 0
                    If tblCal.ListColumns.Count >= 3 Then
                        ja = Val(Replace(CStr(tblCal.DataBodyRange(cr, 3).Value), ",", "."))
                    End If
                    m_CalCap("CAL|" & mDate & "|JO") = jo
                    m_CalCap("CAL|" & mDate & "|JA") = ja
                Next cr
            End If
        End If
    End If
    
    ' Charger PLANT par usine
    Set wsPl = ThisWorkbook.Sheets(SH_PLANT)
    If Not wsPl Is Nothing Then
        Set tblPl = wsPl.ListObjects("tbl_PLANT")
        If Not tblPl Is Nothing Then
            If Not tblPl.DataBodyRange Is Nothing Then
                For pr = 1 To tblPl.DataBodyRange.Rows.Count
                    pid = Trim(CStr(tblPl.DataBodyRange(pr, 1).Value))
                    cRef = Val(Replace(CStr(tblPl.DataBodyRange(pr, 3).Value), ",", "."))
                    mg = Val(Replace(CStr(tblPl.DataBodyRange(pr, 4).Value), ",", "."))
                    hs = Val(Replace(CStr(tblPl.DataBodyRange(pr, 5).Value), ",", "."))
                    st = Val(Replace(CStr(tblPl.DataBodyRange(pr, 6).Value), ",", "."))
                    m_CalCap("PLANT|" & pid & "|CREF") = cRef
                    m_CalCap("PLANT|" & pid & "|MG") = mg
                    m_CalCap("PLANT|" & pid & "|HS") = hs
                    m_CalCap("PLANT|" & pid & "|ST") = st
                Next pr
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub SauvegarderCapaciteChampsEnMemoire(ByVal plantId As String, ByVal moisStr As String)
    If plantId = "" Or moisStr = "" Then Exit Sub
    If IsNumeric(txt_JoursOuvres.Text) Then m_CalCap("CAL|" & moisStr & "|JO") = Val(Replace(txt_JoursOuvres.Text, ",", "."))
    If IsNumeric(txt_JoursArret.Text) Then m_CalCap("CAL|" & moisStr & "|JA") = Val(Replace(txt_JoursArret.Text, ",", "."))
    If IsNumeric(txt_CapNominale.Text) Then m_CalCap("PLANT|" & plantId & "|CREF") = Val(Replace(txt_CapNominale.Text, ",", "."))
    If IsNumeric(txt_MargeCapPlant.Text) Then m_CalCap("PLANT|" & plantId & "|MG") = Val(Replace(txt_MargeCapPlant.Text, ",", ".")) / 100#
    If IsNumeric(txt_CapMaxHS.Text) Then m_CalCap("PLANT|" & plantId & "|HS") = Val(Replace(txt_CapMaxHS.Text, ",", ".")) / 100#
    If IsNumeric(txt_CapMaxST.Text) Then m_CalCap("PLANT|" & plantId & "|ST") = Val(Replace(txt_CapMaxST.Text, ",", "."))
End Sub

Private Sub AfficherCapacitePourSelection(ByVal plantId As String, ByVal moisStr As String)
    Dim kJo As String: kJo = "CAL|" & moisStr & "|JO"
    Dim kJa As String: kJa = "CAL|" & moisStr & "|JA"
    If m_CalCap.Exists(kJo) Then txt_JoursOuvres.Text = CStr(m_CalCap(kJo)) Else txt_JoursOuvres.Text = "21"
    If m_CalCap.Exists(kJa) Then txt_JoursArret.Text = CStr(m_CalCap(kJa)) Else txt_JoursArret.Text = "0"
    
    Dim kCref As String: kCref = "PLANT|" & plantId & "|CREF"
    Dim kMg As String: kMg = "PLANT|" & plantId & "|MG"
    Dim kHs As String: kHs = "PLANT|" & plantId & "|HS"
    Dim kSt As String: kSt = "PLANT|" & plantId & "|ST"
    
    If m_CalCap.Exists(kCref) Then txt_CapNominale.Text = CStr(m_CalCap(kCref)) Else txt_CapNominale.Text = "5000"
    If m_CalCap.Exists(kMg) Then txt_MargeCapPlant.Text = CStr(Round(CDbl(m_CalCap(kMg)) * 100#, 1)) Else txt_MargeCapPlant.Text = "10"
    If m_CalCap.Exists(kHs) Then txt_CapMaxHS.Text = CStr(Round(CDbl(m_CalCap(kHs)) * 100#, 1)) Else txt_CapMaxHS.Text = "20"
    If m_CalCap.Exists(kSt) Then txt_CapMaxST.Text = CStr(m_CalCap(kSt)) Else txt_CapMaxST.Text = "1000"
    
    CalculerCapaciteEstimee
End Sub

Private Sub cbo_PlantCap_Change()
    If m_Initializing Then Exit Sub
    If m_CurrentPlantCap <> "" And m_CurrentMois <> "" Then
        SauvegarderCapaciteChampsEnMemoire m_CurrentPlantCap, m_CurrentMois
    End If
    m_CurrentPlantCap = cbo_PlantCap.Text
    AfficherCapacitePourSelection m_CurrentPlantCap, m_CurrentMois
End Sub

Private Sub cbo_MoisCal_Change()
    If m_Initializing Then Exit Sub
    If m_CurrentPlantCap <> "" And m_CurrentMois <> "" Then
        SauvegarderCapaciteChampsEnMemoire m_CurrentPlantCap, m_CurrentMois
    End If
    m_CurrentMois = cbo_MoisCal.Text
    AfficherCapacitePourSelection m_CurrentPlantCap, m_CurrentMois
End Sub

Private Sub txt_JoursOuvres_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub txt_JoursArret_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub txt_CapNominale_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub txt_MargeCapPlant_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub txt_CapMaxHS_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub txt_CapMaxST_Change()
    If m_Initializing Then Exit Sub
    CalculerCapaciteEstimee
End Sub

Private Sub CalculerCapaciteEstimee()
    Dim jo As Double, ja As Double, cn As Double
    Dim mg As Double, hs As Double, st As Double
    Dim ratioOuvre As Double
    Dim capBase As Double, capHS As Double, capST As Double, capTotale As Double
    
    On Error Resume Next
    jo = Val(Replace(txt_JoursOuvres.Text, ",", "."))
    ja = Val(Replace(txt_JoursArret.Text, ",", "."))
    cn = Val(Replace(txt_CapNominale.Text, ",", "."))
    mg = Val(Replace(txt_MargeCapPlant.Text, ",", ".")) / 100#
    hs = Val(Replace(txt_CapMaxHS.Text, ",", ".")) / 100#
    st = Val(Replace(txt_CapMaxST.Text, ",", "."))
    
    If jo <= 0 Then jo = 21
    If (jo - ja) > 0 Then ratioOuvre = (jo - ja) / jo Else ratioOuvre = 0
    
    capBase = cn * (1# + mg) * ratioOuvre
    capHS = cn * hs
    capST = st
    capTotale = Round(capBase + capHS + capST, 0)
    
    lbl_ValCapDispo.Caption = Format(capTotale, "#,##0") & " unit" & Chr(233) & "s / mois" & _
                              " (Base: " & Format(Round(capBase, 0), "#,##0") & " + HS: " & Format(Round(capHS, 0), "#,##0") & " + ST: " & Format(Round(capST, 0), "#,##0") & ")"
End Sub

' -------------------------------------------------------------
' ACTIONS BOUTONS DU BAS
' -------------------------------------------------------------
Private Sub btn_Fermer_Click()
    Unload Me
End Sub

Private Sub btn_ChargerDefaut_Click()
    If MsgBox("Voulez-vous r" & Chr(233) & "tablir toutes les valeurs par d" & Chr(233) & "faut (Gouvernance, Co" & Chr(251) & "ts, Capacit" & Chr(233) & "s) ?" & vbCrLf & _
              "(Les valeurs actuelles dans le formulaire seront r" & Chr(233) & "initialis" & Chr(233) & "es)", vbQuestion + vbYesNo, TOOL_NAME) = vbNo Then
        Exit Sub
    End If
    
    ' Tab 1
    txt_NiveauService.Text = "95"
    txt_SeuilAttention.Text = "85"
    txt_SeuilSurcharge.Text = "100"
    txt_SeuilDLC.Text = "30"
    txt_StockMax.Text = "50000"
    txt_DelaiLivraison.Text = "5"
    txt_MargeCapacite.Text = "10"
    
    ' Tab 2: Couts de base par usine
    m_Costs("USN-01|Prod") = 150: m_Costs("USN-01|HS") = 350: m_Costs("USN-01|ST") = 500: m_Costs("USN-01|Stockage") = 12: m_Costs("USN-01|Penurie") = 800: m_Costs("USN-01|Peremption") = 200
    m_Costs("USN-02|Prod") = 145: m_Costs("USN-02|HS") = 340: m_Costs("USN-02|ST") = 490: m_Costs("USN-02|Stockage") = 12: m_Costs("USN-02|Penurie") = 800: m_Costs("USN-02|Peremption") = 200
    m_Costs("USN-03|Prod") = 140: m_Costs("USN-03|HS") = 320: m_Costs("USN-03|ST") = 480: m_Costs("USN-03|Stockage") = 11: m_Costs("USN-03|Penurie") = 750: m_Costs("USN-03|Peremption") = 190
    AfficherCoutsPourUsine cbo_CostPlant.Text
    
    ' Tab 3: Capacites de base
    m_CalCap("PLANT|USN-01|CREF") = 5000: m_CalCap("PLANT|USN-01|MG") = 0.1: m_CalCap("PLANT|USN-01|HS") = 0.2: m_CalCap("PLANT|USN-01|ST") = 1000
    m_CalCap("PLANT|USN-02|CREF") = 4000: m_CalCap("PLANT|USN-02|MG") = 0.1: m_CalCap("PLANT|USN-02|HS") = 0.2: m_CalCap("PLANT|USN-02|ST") = 800
    m_CalCap("PLANT|USN-03|CREF") = 2500: m_CalCap("PLANT|USN-03|MG") = 0.1: m_CalCap("PLANT|USN-03|HS") = 0.15: m_CalCap("PLANT|USN-03|ST") = 500
    
    txt_JoursOuvres.Text = "21"
    txt_JoursArret.Text = "0"
    AfficherCapacitePourSelection cbo_PlantCap.Text, cbo_MoisCal.Text
    
    lbl_Info.Caption = "Valeurs par d" & Chr(233) & "faut r" & Chr(233) & "tablies. Cliquez sur 'Enregistrer Param" & Chr(232) & "tres' pour valider."
    lbl_Info.ForeColor = RGB(0, 100, 0)
End Sub

Private Sub btn_Enregistrer_Click()
    Dim okGlob As Boolean
    Dim plants As Variant, cTypes As Variant
    Dim p As Variant, ct As Variant
    Dim kCost As String, pStr As String
    Dim cRef As Double, mg As Double, hs As Double, st As Double
    Dim mIdx As Long, mStr As String
    Dim jo As Long, ja As Long
    Dim rep As VbMsgBoxResult
    
    If Not ValiderTousLesChamps() Then Exit Sub
    
    ' Sauvegarder les saisies actuelles des onglets actifs
    SauvegarderCoutsChampsEnMemoire cbo_CostPlant.Text
    SauvegarderCapaciteChampsEnMemoire cbo_PlantCap.Text, cbo_MoisCal.Text
    
    ' 1. Enregistrer Tab 1 (Parametres Globaux dans tbl_PARAM)
    okGlob = SauvegarderParametresGlobaux(CDbl(txt_NiveauService.Text) / 100#, _
                                         CDbl(txt_SeuilAttention.Text) / 100#, _
                                         CDbl(txt_SeuilSurcharge.Text) / 100#, _
                                         CLng(txt_SeuilDLC.Text), _
                                         CDbl(txt_StockMax.Text), _
                                         CLng(txt_DelaiLivraison.Text), _
                                         CDbl(txt_MargeCapacite.Text) / 100#)
    
    ' 2. Enregistrer Tab 2 (Matrice des Couts dans tbl_COST)
    plants = Array("USN-01", "USN-02", "USN-03")
    cTypes = Array("Prod", "HS", "ST", "Stockage", "Penurie", "Peremption")
    For Each p In plants
        For Each ct In cTypes
            kCost = CStr(p) & "|" & CStr(ct)
            If m_Costs.Exists(kCost) Then
                SauvegarderCoutUsine CStr(p), CStr(ct), CDbl(m_Costs(kCost))
            End If
        Next ct
    Next p
    
    ' 3. Enregistrer Tab 3 (Capacites dans tbl_PLANT et Jours Ouvres dans tbl_CALENDAR)
    For Each p In plants
        pStr = CStr(p)
        cRef = CDbl(m_CalCap("PLANT|" & pStr & "|CREF"))
        mg = CDbl(m_CalCap("PLANT|" & pStr & "|MG"))
        hs = CDbl(m_CalCap("PLANT|" & pStr & "|HS"))
        st = CDbl(m_CalCap("PLANT|" & pStr & "|ST"))
        
        For mIdx = 0 To cbo_MoisCal.ListCount - 1
            mStr = cbo_MoisCal.List(mIdx)
            jo = 21
            ja = 0
            If m_CalCap.Exists("CAL|" & mStr & "|JO") Then jo = CLng(m_CalCap("CAL|" & mStr & "|JO"))
            If m_CalCap.Exists("CAL|" & mStr & "|JA") Then ja = CLng(m_CalCap("CAL|" & mStr & "|JA"))
            SauvegarderCalendrierCapacite mStr, pStr, jo, ja, cRef, mg, hs, st
        Next mIdx
    Next p
    
    LogMsg "PARAMETRES", "INFO", "Enregistrement complet des parametres SOP par " & GetCurrentUser()
    
    ' 4. Message de confirmation et proposition de recalcul
    rep = MsgBox("Param" & Chr(232) & "tres et Structure des Co" & Chr(251) & "ts mis " & Chr(224) & " jour avec succ" & Chr(232) & "s !" & vbCrLf & _
                 "Toutes les modifications ont " & Chr(233) & "t" & Chr(233) & " enregistr" & Chr(233) & "es et trac" & Chr(233) & "es dans l'audit JOURNAL." & vbCrLf & vbCrLf & _
                 "Souhaitez-vous relancer le cycle S&OP pour recalculer les impacts ?", _
                 vbQuestion + vbYesNo, TOOL_NAME)
    
    If rep = vbYes Then
        Unload Me
        Call LancerCycleComplet
    Else
        lbl_Info.Caption = "Param" & Chr(232) & "tres enregistr" & Chr(233) & "s avec succ" & Chr(232) & "s le " & Format(Now, "hh:nn:ss")
        lbl_Info.ForeColor = RGB(0, 120, 0)
    End If
End Sub

Private Function ValiderTousLesChamps() As Boolean
    ValiderTousLesChamps = False
    
    ' Tab 1
    If Not IsNumeric(txt_NiveauService.Text) Or val(txt_NiveauService.Text) <= 0 Or val(txt_NiveauService.Text) > 100 Then
        MsgBox "Le taux de service cible doit " & Chr(234) & "tre compris entre 1% et 100%.", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_NiveauService.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_SeuilAttention.Text) Or val(txt_SeuilAttention.Text) <= 0 Or val(txt_SeuilAttention.Text) > 100 Then
        MsgBox "Le seuil d'alerte Attention doit " & Chr(234) & "tre compris entre 1% et 100%.", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_SeuilAttention.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_SeuilSurcharge.Text) Or val(txt_SeuilSurcharge.Text) < val(txt_SeuilAttention.Text) Then
        MsgBox "Le seuil de surcharge usine doit " & Chr(234) & "tre sup" & Chr(233) & "rieur ou " & Chr(233) & "gal au seuil d'attention.", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_SeuilSurcharge.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_SeuilDLC.Text) Or val(txt_SeuilDLC.Text) <= 0 Then
        MsgBox "Le seuil DLC doit " & Chr(234) & "tre un nombre de jours strictement positif (> 0).", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_SeuilDLC.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_StockMax.Text) Or val(txt_StockMax.Text) <= 0 Then
        MsgBox "Le plafond de stock max doit " & Chr(234) & "tre strictement positif (> 0).", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_StockMax.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_DelaiLivraison.Text) Or val(txt_DelaiLivraison.Text) <= 0 Then
        MsgBox "Le d" & Chr(233) & "lai de r" & Chr(233) & "approvisionnement doit " & Chr(234) & "tre strictement positif.", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_DelaiLivraison.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_MargeCapacite.Text) Or val(txt_MargeCapacite.Text) < 0 Then
        MsgBox "La marge de flexibilit" & Chr(233) & " capacit" & Chr(233) & " doit " & Chr(234) & "tre positive ou nulle.", vbExclamation, TOOL_NAME
        mp_Tabs.Value = 0: txt_MargeCapacite.SetFocus: Exit Function
    End If
    
    ' Tab 2
    If Not IsNumeric(txt_CostProd.Text) Or val(txt_CostProd.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t de production invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostProd.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CostHS.Text) Or val(txt_CostHS.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t des heures suppl" & Chr(233) & "mentaires invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostHS.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CostST.Text) Or val(txt_CostST.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t de sous-traitance invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostST.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CostStockage.Text) Or val(txt_CostStockage.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t de stockage invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostStockage.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CostPenurie.Text) Or val(txt_CostPenurie.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t de p" & Chr(233) & "nurie invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostPenurie.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CostPeremption.Text) Or val(txt_CostPeremption.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t de d" & Chr(233) & "pr" & Chr(233) & "ciation DLC invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 1: txt_CostPeremption.SetFocus: Exit Function
    End If
    
    ' Tab 3
    If Not IsNumeric(txt_JoursOuvres.Text) Or val(txt_JoursOuvres.Text) <= 0 Or val(txt_JoursOuvres.Text) > 31 Then
        MsgBox "Nombre de jours ouvr" & Chr(233) & "s invalide (doit " & Chr(234) & "tre entre 1 et 31).", vbExclamation, TOOL_NAME: mp_Tabs.Value = 2: txt_JoursOuvres.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_JoursArret.Text) Or val(txt_JoursArret.Text) < 0 Or val(txt_JoursArret.Text) > val(txt_JoursOuvres.Text) Then
        MsgBox "Nombre de jours d'arr" & Chr(234) & "t invalide (doit " & Chr(234) & "tre <= jours ouvr" & Chr(233) & "s).", vbExclamation, TOOL_NAME: mp_Tabs.Value = 2: txt_JoursArret.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_CapNominale.Text) Or val(txt_CapNominale.Text) <= 0 Then
        MsgBox "Capacit" & Chr(233) & " nominale invalide (> 0).", vbExclamation, TOOL_NAME: mp_Tabs.Value = 2: txt_CapNominale.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_MargeCapPlant.Text) Or val(txt_MargeCapPlant.Text) < 0 Then
        MsgBox "Marge usine invalide.", vbExclamation, TOOL_NAME: mp_Tabs.Value = 2: txt_MargeCapPlant.SetFocus: Exit Function
    End If
    
    ValiderTousLesChamps = True
End Function


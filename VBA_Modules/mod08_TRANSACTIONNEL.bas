Attribute VB_Name = "mod08_TRANSACTIONNEL"
'==============================================================
' MODULE 08 - VOLET TRANSACTIONNEL (CRUD USERFORMS)
' S&OP Decision Support System
' EMINES - Specifications Academiques S&OP
' Features:
'   T1: Gestion des Produits & Master Data (frm_Produits)
'   T2: Commandes Clients (frm_Commandes)
'   T3: Mouvements & Stocks avec Controle Negatif (frm_MouvementsStock)
'   T5: Ajustements Commerciaux avec Tracabilite (frm_AjustementCommercial)
'==============================================================
Option Explicit

' -------------------------------------------------------------
' 1. MACROS D'OUVERTURE DES USERFORMS (Lancement depuis l'Accueil)
' -------------------------------------------------------------
Public Sub OuvrirGestionProduits()
    On Error GoTo ErrHandler
    InitialiserTablesTransactionnelles
    frm_Produits.Show
    Exit Sub
ErrHandler:
    If Application.UserControl Then MsgBox "Erreur lors de l'ouverture de la Gestion Produits: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Public Sub OuvrirCommandesClients()
    On Error GoTo ErrHandler
    InitialiserTablesTransactionnelles
    frm_Commandes.Show
    Exit Sub
ErrHandler:
    If Application.UserControl Then MsgBox "Erreur lors de l'ouverture des Commandes Clients: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Public Sub OuvrirMouvementsStock()
    On Error GoTo ErrHandler
    InitialiserTablesTransactionnelles
    frm_MouvementsStock.Show
    Exit Sub
ErrHandler:
    If Application.UserControl Then MsgBox "Erreur lors de l'ouverture des Mouvements de Stock: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Public Sub OuvrirAjustementsCommerciaux()
    On Error GoTo ErrHandler
    Dim wsDp As Worksheet
    On Error Resume Next
    Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    On Error GoTo ErrHandler
    If wsDp Is Nothing Or wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Le Demand Plan n'est pas encore genere. Veuillez d'abord lancer l'etape Forecast & Demand Plan.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    frm_AjustementCommercial.Show
    Exit Sub
ErrHandler:
    If Application.UserControl Then MsgBox "Erreur lors de l'ouverture de l'Ajustement Commercial: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Public Sub OuvrirGestionParametres()
    On Error GoTo ErrHandler
    InitialiserTablesTransactionnelles
    frm_Parametres.Show
    Exit Sub
ErrHandler:
    If Application.UserControl Then MsgBox "Erreur lors de l'ouverture des Param" & Chr(232) & "tres S&OP: " & Err.Description, vbCritical, TOOL_NAME
End Sub


Private Function ObtenirOuCreerFeuille(ByVal nomFeuille As String, ByVal tabCouleur As Long) As Worksheet
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nomFeuille
        If tabCouleur <> 0 Then ws.Tab.Color = tabCouleur
    End If
    Set ObtenirOuCreerFeuille = ws
End Function

' -------------------------------------------------------------
' 2. INITIALISATION ET VERIFICATION DES TABLES TRANSACTIONNELLES
' -------------------------------------------------------------
Public Sub InitialiserTablesTransactionnelles()
    On Error Resume Next
    Dim wsProd As Worksheet: Set wsProd = Nothing
    Dim wsCmd As Worksheet: Set wsCmd = Nothing
    Dim wsMvt As Worksheet: Set wsMvt = Nothing
    Dim loProd As ListObject: Set loProd = Nothing
    Dim loCmd As ListObject: Set loCmd = Nothing
    Dim loMvt As ListObject: Set loMvt = Nothing
    
    ' A. Verifier / Completer tbl_PRODUCT
    Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    If Not wsProd Is Nothing Then
        Set loProd = wsProd.ListObjects("tbl_PRODUCT")
        If Not loProd Is Nothing Then
            If loProd.ListColumns.Count < 10 Then
                Dim cNames As Variant
                cNames = Array("Prix_vente", "Cout_standard", "Statut_Actif")
                Dim cn As Variant
                For Each cn In cNames
                    Dim colExists As Boolean: colExists = False
                    Dim lc As ListColumn
                    For Each lc In loProd.ListColumns
                        If UCase(lc.Name) = UCase(CStr(cn)) Then colExists = True: Exit For
                    Next lc
                    If Not colExists Then
                        Set lc = loProd.ListColumns.Add
                        lc.Name = CStr(cn)
                    End If
                Next cn
                
                Dim r As Long
                For r = 1 To loProd.DataBodyRange.Rows.Count
                    If IsEmpty(loProd.DataBodyRange(r, 8).Value) Or loProd.DataBodyRange(r, 8).Value = "" Then loProd.DataBodyRange(r, 8).Value = 250
                    If IsEmpty(loProd.DataBodyRange(r, 9).Value) Or loProd.DataBodyRange(r, 9).Value = "" Then loProd.DataBodyRange(r, 9).Value = 150
                    If IsEmpty(loProd.DataBodyRange(r, 10).Value) Or loProd.DataBodyRange(r, 10).Value = "" Then loProd.DataBodyRange(r, 10).Value = "OUI"
                Next r
            End If
        End If
    End If
    
    ' B. Verifier / Creer feuille COMMANDES
    Set wsCmd = ObtenirOuCreerFeuille("COMMANDES", RGB(0, 112, 192))
    If Not wsCmd Is Nothing Then
        Set loCmd = Nothing
        On Error Resume Next: Set loCmd = wsCmd.ListObjects("tbl_COMMANDES"): On Error GoTo 0
        If loCmd Is Nothing Then
            wsCmd.Cells.Clear
            wsCmd.Cells(1, 1).Value = "Commande_ID"
            wsCmd.Cells(1, 2).Value = "Client"
            wsCmd.Cells(1, 3).Value = "SKU_ID"
            wsCmd.Cells(1, 4).Value = "Quantite"
            wsCmd.Cells(1, 5).Value = "Date_Commande"
            wsCmd.Cells(1, 6).Value = "Date_Livraison_Prevue"
            wsCmd.Cells(1, 7).Value = "Date_Livraison_Reelle"
            wsCmd.Cells(1, 8).Value = "Statut"
            wsCmd.Cells(1, 9).Value = "Commentaire"
            
            Dim initCmds As Variant
            initCmds = Array( _
                Array("CMD-2023-001", "Carrefour Market", "SKU-01", 1200, "2023-05-15", "2023-05-20", "2023-05-20", "Livr" & Chr(233) & "e", "Livraison conforme"), _
                Array("CMD-2023-002", "Marjane Holding", "SKU-02", 950, "2023-05-18", "2023-05-24", "2023-05-25", "Livr" & Chr(233) & "e en retard", "Retard transporteur 24h"), _
                Array("CMD-2023-003", "Grossiste Casablanca", "SKU-03", 2000, "2023-06-01", "2023-06-08", "", "En cours", "En preparation entrepot"), _
                Array("CMD-2023-004", "Export Dakar", "SKU-05", 500, "2023-06-02", "2023-06-15", "", "En cours", "Conteneur maritime reserve"), _
                Array("CMD-2023-005", "Acima Rabat", "SKU-04", 800, "2023-05-28", "2023-06-02", "", "En retard", "Rupture stock local") _
            )
            Dim ic As Long
            For ic = 0 To UBound(initCmds)
                Dim colIdx As Long
                For colIdx = 0 To 8
                    wsCmd.Cells(ic + 2, colIdx + 1).Value = initCmds(ic)(colIdx)
                Next colIdx
            Next ic
            
            Dim cmdHdrs As Variant
            cmdHdrs = Array("Commande_ID", "Client", "SKU_ID", "Quantite", "Date_Commande", "Date_Livraison_Prevue", "Date_Livraison_Reelle", "Statut", "Commentaire")
            SetupTable wsCmd, "tbl_COMMANDES", cmdHdrs, UBound(initCmds) + 3, "TableStyleMedium2"
        End If
    End If
    
    ' C. Verifier / Creer feuille MOUVEMENTS
    Set wsMvt = ObtenirOuCreerFeuille("MOUVEMENTS", RGB(0, 176, 80))
    If Not wsMvt Is Nothing Then
        Set loMvt = Nothing
        On Error Resume Next: Set loMvt = wsMvt.ListObjects("tbl_MOUVEMENTS"): On Error GoTo 0
        If loMvt Is Nothing Then
            wsMvt.Cells.Clear
            wsMvt.Cells(1, 1).Value = "Mvt_ID"
            wsMvt.Cells(1, 2).Value = "Date_Mvt"
            wsMvt.Cells(1, 3).Value = "SKU_ID"
            wsMvt.Cells(1, 4).Value = "Plant_ID"
            wsMvt.Cells(1, 5).Value = "Type_Mvt"
            wsMvt.Cells(1, 6).Value = "Quantite"
            wsMvt.Cells(1, 7).Value = "Stock_Avant"
            wsMvt.Cells(1, 8).Value = "Stock_Apres"
            wsMvt.Cells(1, 9).Value = "Motif"
            wsMvt.Cells(1, 10).Value = "Utilisateur"
            
            Dim initMvts As Variant
            initMvts = Array( _
                Array("MVT-20230501-01", "2023-05-01 08:30:00", "SKU-01", "USN-01", "Entr" & Chr(233) & "e R" & Chr(233) & "ception (+)", 1500, 3500, 5000, "Production Lot N-401", "ADMIN"), _
                Array("MVT-20230505-02", "2023-05-05 14:15:00", "SKU-01", "USN-01", "Sortie Livraison (-)", 800, 5000, 4200, "Livraison Carrefour CMD-001", "LOGISTIQUE"), _
                Array("MVT-20230510-03", "2023-05-10 10:00:00", "SKU-02", "USN-01", "Entr" & Chr(233) & "e R" & Chr(233) & "ception (+)", 1000, 3000, 4000, "Production Lot C-102", "PROD_CHEF"), _
                Array("MVT-20230515-04", "2023-05-15 16:45:00", "SKU-03", "USN-02", "Perte DLC / Rebut (-)", 50, 3550, 3500, "Endommagement emballage transport", "QUALITE") _
            )
            Dim im As Long
            For im = 0 To UBound(initMvts)
                For colIdx = 0 To 9
                    wsMvt.Cells(im + 2, colIdx + 1).Value = initMvts(im)(colIdx)
                Next colIdx
            Next im
            
            Dim mvtHdrs As Variant
            mvtHdrs = Array("Mvt_ID", "Date_Mvt", "SKU_ID", "Plant_ID", "Type_Mvt", "Quantite", "Stock_Avant", "Stock_Apres", "Motif", "Utilisateur")
            SetupTable wsMvt, "tbl_MOUVEMENTS", mvtHdrs, UBound(initMvts) + 3, "TableStyleMedium7"
        End If
    End If
End Sub

' -------------------------------------------------------------
' 3. GESTION DU STOCK PHYSIQUE & REGLE DE NON-NEGATIVITE (T3)
' -------------------------------------------------------------
Public Function GetStockActuel(ByVal sku As String) As Double
    On Error Resume Next
    Dim baseStock As Double: baseStock = 1000
    
    Dim pVal As Variant
    pVal = GetParam("STOCK_INIT_" & Replace(sku, "-", ""))
    If IsNumeric(pVal) Then
        baseStock = CDbl(pVal)
    Else
        Select Case sku
            Case "SKU-01": baseStock = 5000
            Case "SKU-02": baseStock = 4000
            Case "SKU-03": baseStock = 3500
            Case "SKU-04": baseStock = 3000
            Case "SKU-05": baseStock = 2000
        End Select
    End If
    
    Dim wsMvt As Worksheet: Set wsMvt = Nothing
    Set wsMvt = ThisWorkbook.Sheets("MOUVEMENTS")
    If Not wsMvt Is Nothing Then
        Dim loMvt As ListObject
        Set loMvt = wsMvt.ListObjects("tbl_MOUVEMENTS")
        If Not loMvt Is Nothing And Not loMvt.DataBodyRange Is Nothing Then
            Dim r As Long
            Dim deltaNet As Double: deltaNet = 0
            For r = 1 To loMvt.DataBodyRange.Rows.Count
                If CStr(loMvt.DataBodyRange(r, 3).Value) = sku Then
                    Dim tMvt As String: tMvt = CStr(loMvt.DataBodyRange(r, 5).Value)
                    Dim q As Double: q = CDbl(loMvt.DataBodyRange(r, 6).Value)
                    If InStr(1, tMvt, "(+)") > 0 Or InStr(1, tMvt, "Entr", vbTextCompare) > 0 Then
                        deltaNet = deltaNet + q
                    ElseIf InStr(1, tMvt, "(-)") > 0 Or InStr(1, tMvt, "Sortie", vbTextCompare) > 0 Or InStr(1, tMvt, "Rebut", vbTextCompare) > 0 Then
                        deltaNet = deltaNet - q
                    End If
                End If
            Next r
            GetStockActuel = baseStock + deltaNet
            Exit Function
        End If
    End If
    GetStockActuel = baseStock
End Function

Public Function EnregistrerMouvementStock(ByVal sku As String, ByVal plant As String, ByVal typeMvt As String, ByVal qte As Double, ByVal motif As String) As Boolean
    On Error GoTo ErrMvt
    If qte <= 0 Then
        If Application.UserControl Then MsgBox "La quantit" & Chr(233) & " du mouvement doit " & Chr(234) & "tre strictement positive.", vbExclamation, TOOL_NAME
        EnregistrerMouvementStock = False
        Exit Function
    End If
    If Trim(motif) = "" Then
        If Application.UserControl Then MsgBox "Le motif du mouvement est OBLIGATOIRE pour la tra" & Chr(231) & "abilit" & Chr(233) & ".", vbExclamation, TOOL_NAME
        EnregistrerMouvementStock = False
        Exit Function
    End If
    
    Dim stockAvant As Double
    stockAvant = GetStockActuel(sku)
    
    Dim delta As Double
    If InStr(1, typeMvt, "(+)") > 0 Or InStr(1, typeMvt, "Entr", vbTextCompare) > 0 Then
        delta = qte
    ElseIf InStr(1, typeMvt, "(-)") > 0 Or InStr(1, typeMvt, "Sortie", vbTextCompare) > 0 Or InStr(1, typeMvt, "Rebut", vbTextCompare) > 0 Then
        delta = -qte
    Else
        delta = -qte
    End If
    
    ' REGLE CRITIQUE DE NON-NEGATIVITE DU STOCK
    If (stockAvant + delta) < 0 Then
        If Application.UserControl Then
            MsgBox "TRANSACTION BLOQU" & Chr(201) & "E !" & vbCrLf & vbCrLf & _
                   "R" & Chr(232) & "gle de non-n" & Chr(233) & "gativit" & Chr(233) & " des stocks viol" & Chr(233) & "e :" & vbCrLf & _
                   "- Stock physique actuel : " & Format(stockAvant, "#,##0") & " unit" & Chr(233) & "s" & vbCrLf & _
                   "- Quantit" & Chr(233) & " demand" & Chr(233) & "e en sortie : " & Format(qte, "#,##0") & " unit" & Chr(233) & "s" & vbCrLf & _
                   "- Stock r" & Chr(233) & "sultant impossible : " & Format(stockAvant + delta, "#,##0") & " unit" & Chr(233) & "s" & vbCrLf & vbCrLf & _
                   "L'op" & Chr(233) & "ration a " & Chr(233) & "t" & Chr(233) & " annul" & Chr(233) & "e sans modifier les tables.", vbCritical, "Contr" & Chr(244) & "le des Stocks"
        End If
        EnregistrerMouvementStock = False
        Exit Function
    End If
    
    Dim stockApres As Double: stockApres = stockAvant + delta
    
    Dim wsMvt As Worksheet: Set wsMvt = ThisWorkbook.Sheets("MOUVEMENTS")
    Dim loMvt As ListObject: Set loMvt = wsMvt.ListObjects("tbl_MOUVEMENTS")
    Dim newRow As ListRow: Set newRow = loMvt.ListRows.Add
    
    Dim mvtId As String: mvtId = "MVT-" & Format(Now, "YYYYMMDD-HHNNSS")
    newRow.Range(1, 1).Value = mvtId
    newRow.Range(1, 2).Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    newRow.Range(1, 3).Value = sku
    newRow.Range(1, 4).Value = plant
    newRow.Range(1, 5).Value = typeMvt
    newRow.Range(1, 6).Value = qte
    newRow.Range(1, 7).Value = stockAvant
    newRow.Range(1, 8).Value = stockApres
    newRow.Range(1, 9).Value = motif
    newRow.Range(1, 10).Value = GetCurrentUser()
    
    LogMsg "STOCK", "INFO", "Mvt " & typeMvt & " " & sku & " Qte=" & qte & " [Stock: " & stockAvant & " -> " & stockApres & "] Motif=" & motif
    EnregistrerMouvementStock = True
    Exit Function
ErrMvt:
    LogMsg "STOCK", "ERROR", "Erreur Mouvement Stock: " & Err.Description
    If Application.UserControl Then MsgBox "Erreur Mouvement Stock: " & Err.Description, vbCritical, TOOL_NAME
    EnregistrerMouvementStock = False
End Function

' -------------------------------------------------------------
' 4. GESTION DES PRODUITS (T1)
' -------------------------------------------------------------
Public Function SauvegarderProduit(ByVal sku As String, ByVal nom As String, ByVal famille As String, ByVal plant As String, ByVal poids As Double, ByVal dlc As Integer, ByVal prix As Double, ByVal cout As Double, ByVal actif As Boolean, ByVal isNew As Boolean) As Boolean
    On Error GoTo ErrProd
    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    Dim loProd As ListObject: Set loProd = wsProd.ListObjects("tbl_PRODUCT")
    
    If Trim(sku) = "" Or Trim(nom) = "" Or Trim(famille) = "" Or Trim(plant) = "" Then
        If Application.UserControl Then MsgBox "Veuillez renseigner tous les champs obligatoires (Code SKU, Nom, Famille, Usine).", vbExclamation, TOOL_NAME
        SauvegarderProduit = False
        Exit Function
    End If
    
    Dim r As Long
    Dim foundRow As Long: foundRow = 0
    For r = 1 To loProd.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(loProd.DataBodyRange(r, 1).Value))) = UCase(Trim(sku)) Then
            foundRow = r
            Exit For
        End If
    Next r
    
    If isNew Then
        If foundRow > 0 Then
            If Application.UserControl Then MsgBox "Le code SKU '" & sku & "' existe d" & Chr(233) & "j" & Chr(224) & " dans le catalogue !", vbExclamation, TOOL_NAME
            SauvegarderProduit = False
            Exit Function
        End If
        Dim newRow As ListRow: Set newRow = loProd.ListRows.Add
        foundRow = newRow.Index
    Else
        If foundRow = 0 Then
            If Application.UserControl Then MsgBox "Produit '" & sku & "' introuvable pour modification.", vbCritical, TOOL_NAME
            SauvegarderProduit = False
            Exit Function
        End If
    End If
    
    loProd.DataBodyRange(foundRow, 1).Value = sku
    loProd.DataBodyRange(foundRow, 2).Value = nom
    loProd.DataBodyRange(foundRow, 3).Value = famille
    loProd.DataBodyRange(foundRow, 4).Value = plant
    loProd.DataBodyRange(foundRow, 5).Value = poids
    loProd.DataBodyRange(foundRow, 6).Value = dlc
    If loProd.ListColumns.Count >= 7 And IsEmpty(loProd.DataBodyRange(foundRow, 7).Value) Then
        loProd.DataBodyRange(foundRow, 7).Value = "SS_NORMAL"
    End If
    If loProd.ListColumns.Count >= 8 Then loProd.DataBodyRange(foundRow, 8).Value = prix
    If loProd.ListColumns.Count >= 9 Then loProd.DataBodyRange(foundRow, 9).Value = cout
    If loProd.ListColumns.Count >= 10 Then
        If actif Then loProd.DataBodyRange(foundRow, 10).Value = "OUI" Else loProd.DataBodyRange(foundRow, 10).Value = "NON"
    End If
    
    Dim actStr As String: If isNew Then actStr = "Ajout" Else actStr = "Modification"
    LogMsg "PRODUIT", "INFO", actStr & " produit " & sku & " (" & nom & ") Usine=" & plant
    SauvegarderProduit = True
    Exit Function
ErrProd:
    LogMsg "PRODUIT", "ERROR", "Erreur Produit: " & Err.Description
    If Application.UserControl Then MsgBox "Erreur Sauvegarde Produit: " & Err.Description, vbCritical, TOOL_NAME
    SauvegarderProduit = False
End Function

Public Function DesactiverProduit(ByVal sku As String) As Boolean
    On Error GoTo ErrDesact
    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    Dim loProd As ListObject: Set loProd = wsProd.ListObjects("tbl_PRODUCT")
    
    Dim r As Long
    For r = 1 To loProd.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(loProd.DataBodyRange(r, 1).Value))) = UCase(Trim(sku)) Then
            If loProd.ListColumns.Count >= 10 Then
                loProd.DataBodyRange(r, 10).Value = "NON"
            End If
            LogMsg "PRODUIT", "INFO", "D" & Chr(233) & "sactivation produit " & sku
            DesactiverProduit = True
            Exit Function
        End If
    Next r
    If Application.UserControl Then MsgBox "Produit introuvable.", vbExclamation, TOOL_NAME
    DesactiverProduit = False
    Exit Function
ErrDesact:
    If Application.UserControl Then MsgBox "Erreur d" & Chr(233) & "sactivation: " & Err.Description, vbCritical, TOOL_NAME
    DesactiverProduit = False
End Function

' -------------------------------------------------------------
' 5. GESTION DES COMMANDES CLIENTS (T2)
' -------------------------------------------------------------
Public Function SauvegarderCommande(ByVal cmdId As String, ByVal client As String, ByVal sku As String, ByVal qte As Double, ByVal dateCmd As String, ByVal datePrev As String, ByVal dateReel As String, ByVal statut As String, ByVal comm As String, ByVal isNew As Boolean) As Boolean
    On Error GoTo ErrCmd
    Dim wsCmd As Worksheet: Set wsCmd = ThisWorkbook.Sheets("COMMANDES")
    Dim loCmd As ListObject: Set loCmd = wsCmd.ListObjects("tbl_COMMANDES")
    
    If Trim(cmdId) = "" Or Trim(client) = "" Or Trim(sku) = "" Or qte <= 0 Then
        If Application.UserControl Then MsgBox "Champs obligatoires manquants ou quantit" & Chr(233) & " invalide.", vbExclamation, TOOL_NAME
        SauvegarderCommande = False
        Exit Function
    End If
    
    Dim r As Long
    Dim foundRow As Long: foundRow = 0
    For r = 1 To loCmd.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(loCmd.DataBodyRange(r, 1).Value))) = UCase(Trim(cmdId)) Then
            foundRow = r
            Exit For
        End If
    Next r
    
    If isNew Then
        If foundRow > 0 Then
            If Application.UserControl Then MsgBox "La commande '" & cmdId & "' existe d" & Chr(233) & "j" & Chr(224) & " !", vbExclamation, TOOL_NAME
            SauvegarderCommande = False
            Exit Function
        End If
        Dim newRow As ListRow: Set newRow = loCmd.ListRows.Add
        foundRow = newRow.Index
    Else
        If foundRow = 0 Then
            If Application.UserControl Then MsgBox "Commande introuvable pour mise " & Chr(224) & " jour.", vbCritical, TOOL_NAME
            SauvegarderCommande = False
            Exit Function
        End If
    End If
    
    loCmd.DataBodyRange(foundRow, 1).Value = cmdId
    loCmd.DataBodyRange(foundRow, 2).Value = client
    loCmd.DataBodyRange(foundRow, 3).Value = sku
    loCmd.DataBodyRange(foundRow, 4).Value = qte
    loCmd.DataBodyRange(foundRow, 5).Value = dateCmd
    loCmd.DataBodyRange(foundRow, 6).Value = datePrev
    loCmd.DataBodyRange(foundRow, 7).Value = dateReel
    loCmd.DataBodyRange(foundRow, 8).Value = statut
    loCmd.DataBodyRange(foundRow, 9).Value = comm
    
    Dim actStr As String: If isNew Then actStr = "Cr" & Chr(233) & "ation" Else actStr = "Mise " & Chr(224) & " jour"
    LogMsg "COMMANDE", "INFO", actStr & " commande " & cmdId & " Client=" & client & " SKU=" & sku & " Qte=" & qte & " Statut=" & statut
    SauvegarderCommande = True
    Exit Function
ErrCmd:
    If Application.UserControl Then MsgBox "Erreur Sauvegarde Commande: " & Err.Description, vbCritical, TOOL_NAME
    SauvegarderCommande = False
End Function

Public Function AnnulerCommande(ByVal cmdId As String) As Boolean
    On Error GoTo ErrAnnul
    Dim wsCmd As Worksheet: Set wsCmd = ThisWorkbook.Sheets("COMMANDES")
    Dim loCmd As ListObject: Set loCmd = wsCmd.ListObjects("tbl_COMMANDES")
    
    Dim r As Long
    For r = 1 To loCmd.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(loCmd.DataBodyRange(r, 1).Value))) = UCase(Trim(cmdId)) Then
            loCmd.DataBodyRange(r, 8).Value = "Annul" & Chr(233) & "e"
            loCmd.DataBodyRange(r, 9).Value = loCmd.DataBodyRange(r, 9).Value & " [Annul" & Chr(233) & "e le " & Format(Now, "YYYY-MM-DD") & " par " & GetCurrentUser() & "]"
            LogMsg "COMMANDE", "WARNING", "Annulation commande " & cmdId
            AnnulerCommande = True
            Exit Function
        End If
    Next r
    If Application.UserControl Then MsgBox "Commande non trouv" & Chr(233) & "e.", vbExclamation, TOOL_NAME
    AnnulerCommande = False
    Exit Function
ErrAnnul:
    If Application.UserControl Then MsgBox "Erreur annulation: " & Err.Description, vbCritical, TOOL_NAME
    AnnulerCommande = False
End Function

' -------------------------------------------------------------
' 6. AJUSTEMENT COMMERCIAL DE LA DEMANDE (T5)
' -------------------------------------------------------------
Public Function AppliquerAjustementCommercial(ByVal sku As String, ByVal moisStr As String, ByVal typeAjust As String, ByVal valeur As Double, ByVal motif As String) As Boolean
    On Error GoTo ErrAjust
    If Trim(motif) = "" Then
        If Application.UserControl Then
            MsgBox "Le motif de l'ajustement est OBLIGATOIRE pour la tra" & Chr(231) & "abilit" & Chr(233) & " S&OP !" & vbCrLf & _
                   "(Exemple: 'Promotion rentr" & Chr(233) & "e scolaire', 'Gain appel d'offres', 'P" & Chr(233) & "nurie concurrent')", vbExclamation, "Validation S&OP"
        End If
        AppliquerAjustementCommercial = False
        Exit Function
    End If
    
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim loDp As ListObject: Set loDp = wsDp.ListObjects("tbl_DEMAND_PLAN")
    
    If loDp Is Nothing Or loDp.DataBodyRange Is Nothing Then
        If Application.UserControl Then MsgBox "Tableau DEMAND_PLAN introuvable ou vide.", vbCritical, TOOL_NAME
        AppliquerAjustementCommercial = False
        Exit Function
    End If
    
    Dim targetRow As Long: targetRow = 0
    Dim r As Long
    For r = 1 To loDp.DataBodyRange.Rows.Count
        Dim rSku As String: rSku = CStr(loDp.DataBodyRange(r, 1).Value)
        Dim rMois As String: rMois = Format(loDp.DataBodyRange(r, 2).Value, "YYYY-MM")
        If UCase(Trim(rSku)) = UCase(Trim(sku)) And rMois = moisStr Then
            targetRow = r
            Exit For
        End If
    Next r
    
    If targetRow = 0 Then
        If Application.UserControl Then MsgBox "Ligne introuvable pour SKU " & sku & " et mois " & moisStr & " dans le Demand Plan.", vbCritical, TOOL_NAME
        AppliquerAjustementCommercial = False
        Exit Function
    End If
    
    Dim qteStat As Double: qteStat = CDbl(loDp.DataBodyRange(targetRow, 3).Value)
    Dim ajustFinal As Double
    
    If InStr(1, typeAjust, "Pourcentage", vbTextCompare) > 0 Or InStr(1, typeAjust, "%", vbTextCompare) > 0 Then
        ajustFinal = Round(qteStat * (valeur / 100#), 0)
    Else
        ajustFinal = Round(valeur, 0)
    End If
    
    loDp.DataBodyRange(targetRow, 4).Value = ajustFinal
    loDp.DataBodyRange(targetRow, 7).Value = motif & " [" & GetCurrentUser() & " " & Format(Now, "dd/mm hh:nn") & "]"
    
    Dim rowSheet As Long: rowSheet = loDp.DataBodyRange(targetRow, 1).Row
    loDp.DataBodyRange(targetRow, 5).Formula = "=C" & rowSheet & "+D" & rowSheet
    
    Dim qteFinale As Double: qteFinale = qteStat + ajustFinal
    LogMsg "DEMANDPLAN", "INFO", "Ajustement commercial: " & sku & " " & moisStr & " : Stat=" & qteStat & " Ajust=" & ajustFinal & " Final=" & qteFinale & " Motif=" & motif
    
    AppliquerAjustementCommercial = True
    Exit Function
ErrAjust:
    LogMsg "DEMANDPLAN", "ERROR", "Erreur Ajustement Commercial: " & Err.Description
    If Application.UserControl Then MsgBox "Erreur Ajustement Commercial: " & Err.Description, vbCritical, TOOL_NAME
    AppliquerAjustementCommercial = False
End Function

' -------------------------------------------------------------
' 7. GESTION DES PARAMETRES GLOBAUX & STRUCTURE DES COUTS (T1)
' -------------------------------------------------------------
Public Sub EnregistrerJournalAudit(ByVal section As String, ByVal paramNom As String, ByVal valOld As String, ByVal valNew As String, Optional ByVal commentaire As String = "")
    On Error Resume Next
    Dim wsJ As Worksheet
    Set wsJ = ObtenirOuCreerFeuille("JOURNAL", RGB(112, 48, 160))
    If wsJ Is Nothing Then Exit Sub
    
    Dim loJ As ListObject
    Set loJ = Nothing
    On Error Resume Next: Set loJ = wsJ.ListObjects("tbl_JOURNAL"): On Error GoTo 0
    
    If loJ Is Nothing Then
        wsJ.Cells.Clear
        wsJ.Cells(1, 1).Value = "ID"
        wsJ.Cells(1, 2).Value = "Date_Heure"
        wsJ.Cells(1, 3).Value = "Utilisateur"
        wsJ.Cells(1, 4).Value = "Section"
        wsJ.Cells(1, 5).Value = "Parametre"
        wsJ.Cells(1, 6).Value = "Valeur_Precedente"
        wsJ.Cells(1, 7).Value = "Nouvelle_Valeur"
        wsJ.Cells(1, 8).Value = "Commentaire"
        
        Dim jHdrs As Variant
        jHdrs = Array("ID", "Date_Heure", "Utilisateur", "Section", "Parametre", "Valeur_Precedente", "Nouvelle_Valeur", "Commentaire")
        SetupTable wsJ, "tbl_JOURNAL", jHdrs, 2, "TableStyleMedium6"
        Set loJ = wsJ.ListObjects("tbl_JOURNAL")
    End If
    
    If Not loJ Is Nothing Then
        Dim newRow As ListRow
        Set newRow = loJ.ListRows.Add
        Dim nextID As Long: nextID = loJ.DataBodyRange.Rows.Count
        newRow.Range(1, 1).Value = "AUD-" & Format(Now, "YYYYMMDD") & "-" & Format(nextID, "000")
        newRow.Range(1, 2).Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
        newRow.Range(1, 3).Value = GetCurrentUser()
        newRow.Range(1, 4).Value = section
        newRow.Range(1, 5).Value = paramNom
        newRow.Range(1, 6).Value = valOld
        newRow.Range(1, 7).Value = valNew
        newRow.Range(1, 8).Value = commentaire
    End If
    
    LogMsg "AUDIT", "INFO", "[" & section & "] " & paramNom & " : " & valOld & " -> " & valNew & " (" & commentaire & ")"
End Sub

Public Function SauvegarderParametresGlobaux(ByVal niveauService As Double, _
                                             ByVal seuilAttention As Double, _
                                             ByVal seuilSurcharge As Double, _
                                             ByVal seuilDLC As Long, _
                                             ByVal stockMax As Double, _
                                             ByVal delaiLivraison As Long, _
                                             ByVal margeCapacite As Double) As Boolean
    On Error GoTo ErrGlob
    Dim oldNS As String: oldNS = CStr(GetParam("NIVEAU_SERVICE"))
    Dim oldAtt As String: oldAtt = CStr(GetParam("SEUIL_ALERTE_GAP"))
    Dim oldSur As String: oldSur = CStr(GetParam("SEUIL_SURCHARGE_CAP"))
    Dim oldDLC As String: oldDLC = CStr(GetParam("SEUIL_EXPIRATION_DLC"))
    Dim oldMax As String: oldMax = CStr(GetParam("STOCK_MAX_GLOBAL"))
    Dim oldDel As String: oldDel = CStr(GetParam("DELAI_LIVRAISON"))
    Dim oldMar As String: oldMar = CStr(GetParam("MARGE_CAPACITE"))
    
    SetParam "NIVEAU_SERVICE", niveauService, "%", "Parametre", "Niveau de service cible (ex: 0.95 = 95%)"
    SetParam "SEUIL_ALERTE_GAP", seuilAttention, "%", "Parametre", "Seuil d'alerte utilisation capacite (Attention/Orange)"
    SetParam "SEUIL_SURCHARGE_CAP", seuilSurcharge, "%", "Parametre", "Seuil de surcharge critique capacite (Rouge)"
    SetParam "SEUIL_EXPIRATION_DLC", seuilDLC, "jours", "Parametre", "Seuil critique d'expiration DLC"
    SetParam "STOCK_MAX_GLOBAL", stockMax, "unites", "Parametre", "Plafond de stock max global entrepot"
    SetParam "DELAI_LIVRAISON", delaiLivraison, "jours", "Parametre", "Delai de reapprovisionnement moyen"
    SetParam "MARGE_CAPACITE", margeCapacite, "%", "Hypothese", "Marge sur capacite observee"
    
    If oldNS <> CStr(niveauService) Then EnregistrerJournalAudit "PARAMETRES", "NIVEAU_SERVICE", oldNS, CStr(niveauService), "Gouvernance S&OP"
    If oldAtt <> CStr(seuilAttention) Then EnregistrerJournalAudit "PARAMETRES", "SEUIL_ALERTE_GAP", oldAtt, CStr(seuilAttention), "Seuil Attention"
    If oldSur <> CStr(seuilSurcharge) Then EnregistrerJournalAudit "PARAMETRES", "SEUIL_SURCHARGE_CAP", oldSur, CStr(seuilSurcharge), "Seuil Surcharge"
    If oldDLC <> CStr(seuilDLC) Then EnregistrerJournalAudit "PARAMETRES", "SEUIL_EXPIRATION_DLC", oldDLC, CStr(seuilDLC), "Alerte DLC"
    If oldMax <> CStr(stockMax) Then EnregistrerJournalAudit "PARAMETRES", "STOCK_MAX_GLOBAL", oldMax, CStr(stockMax), "Plafond Stock"
    If oldDel <> CStr(delaiLivraison) Then EnregistrerJournalAudit "PARAMETRES", "DELAI_LIVRAISON", oldDel, CStr(delaiLivraison), "Delai Appro"
    If oldMar <> CStr(margeCapacite) Then EnregistrerJournalAudit "PARAMETRES", "MARGE_CAPACITE", oldMar, CStr(margeCapacite), "Marge Reseau"
    
    SauvegarderParametresGlobaux = True
    Exit Function
ErrGlob:
    LogMsg "PARAMETRES", "ERROR", "Erreur SauvegarderParametresGlobaux: " & Err.Description
    SauvegarderParametresGlobaux = False
End Function

Public Function SauvegarderCoutUsine(ByVal plantId As String, ByVal coutType As String, ByVal montant As Double) As Boolean
    On Error GoTo ErrCout
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets(SH_COST)
    If wsC Is Nothing Then Exit Function
    Dim loC As ListObject: Set loC = wsC.ListObjects("tbl_COST")
    If loC Is Nothing Or loC.DataBodyRange Is Nothing Then Exit Function
    
    Dim r As Long
    Dim oldVal As String: oldVal = ""
    For r = 1 To loC.DataBodyRange.Rows.Count
        Dim rType As String: rType = Trim(CStr(loC.DataBodyRange(r, 1).Value))
        Dim rPlant As String: rPlant = Trim(CStr(loC.DataBodyRange(r, 2).Value))
        If UCase(rType) = UCase(Trim(coutType)) And UCase(rPlant) = UCase(Trim(plantId)) Then
            oldVal = CStr(loC.DataBodyRange(r, 3).Value)
            loC.DataBodyRange(r, 3).Value = montant
            If oldVal <> CStr(montant) Then
                EnregistrerJournalAudit "COST", plantId & "_" & coutType, oldVal, CStr(montant), "Mise a jour cout unitaire"
            End If
            SauvegarderCoutUsine = True
            Exit Function
        End If
    Next r
    
    Dim newRow As ListRow
    Set newRow = loC.ListRows.Add
    newRow.Range(1, 1).Value = coutType
    newRow.Range(1, 2).Value = plantId
    newRow.Range(1, 3).Value = montant
    newRow.Range(1, 4).Value = "UM"
    EnregistrerJournalAudit "COST", plantId & "_" & coutType, "N/A", CStr(montant), "Nouveau cout unitaire"
    SauvegarderCoutUsine = True
    Exit Function
ErrCout:
    LogMsg "COST", "ERROR", "Erreur SauvegarderCoutUsine: " & Err.Description
    SauvegarderCoutUsine = False
End Function

Public Function SauvegarderCalendrierCapacite(ByVal moisStr As String, ByVal plantId As String, ByVal joursOuvres As Long, ByVal joursArret As Long, ByVal capNominale As Double, ByVal margePct As Double, ByVal hsPct As Double, ByVal stMax As Double) As Boolean
    On Error GoTo ErrCalCap
    ' 1. Mise a jour CALENDAR
    Dim wsCal As Worksheet: Set wsCal = ThisWorkbook.Sheets(SH_CALENDAR)
    If Not wsCal Is Nothing Then
        Dim tblCal As ListObject: Set tblCal = wsCal.ListObjects("tbl_CALENDAR")
        If Not tblCal Is Nothing And Not tblCal.DataBodyRange Is Nothing Then
            Dim cr As Long
            For cr = 1 To tblCal.DataBodyRange.Rows.Count
                Dim mDate As String: mDate = Format(tblCal.DataBodyRange(cr, 2).Value, "YYYY-MM")
                If mDate = moisStr Then
                    Dim oldJO As String: oldJO = CStr(tblCal.DataBodyRange(cr, 1).Value)
                    tblCal.DataBodyRange(cr, 1).Value = joursOuvres
                    If tblCal.ListColumns.Count >= 3 Then tblCal.DataBodyRange(cr, 3).Value = joursArret
                    If oldJO <> CStr(joursOuvres) Then
                        EnregistrerJournalAudit "CALENDAR", "Jours_ouvres_" & moisStr, oldJO, CStr(joursOuvres), "Ajustement calendrier"
                    End If
                    Exit For
                End If
            Next cr
        End If
    End If
    
    ' 2. Mise a jour PLANT
    Dim wsPl As Worksheet: Set wsPl = ThisWorkbook.Sheets(SH_PLANT)
    If Not wsPl Is Nothing Then
        Dim tblPl As ListObject: Set tblPl = wsPl.ListObjects("tbl_PLANT")
        If Not tblPl Is Nothing And Not tblPl.DataBodyRange Is Nothing Then
            Dim pr As Long
            For pr = 1 To tblPl.DataBodyRange.Rows.Count
                If UCase(Trim(CStr(tblPl.DataBodyRange(pr, 1).Value))) = UCase(Trim(plantId)) Then
                    Dim oldCap As String: oldCap = CStr(tblPl.DataBodyRange(pr, 3).Value)
                    tblPl.DataBodyRange(pr, 3).Value = capNominale
                    tblPl.DataBodyRange(pr, 4).Value = margePct
                    tblPl.DataBodyRange(pr, 5).Value = hsPct
                    tblPl.DataBodyRange(pr, 6).Value = stMax
                    If oldCap <> CStr(capNominale) Then
                        EnregistrerJournalAudit "PLANT", plantId & "_Capacite_ref", oldCap, CStr(capNominale), "Ajustement capacite nominale"
                    End If
                    Exit For
                End If
            Next pr
        End If
    End If
    
    SauvegarderCalendrierCapacite = True
    Exit Function
ErrCalCap:
    LogMsg "CAPACITY", "ERROR", "Erreur SauvegarderCalendrierCapacite: " & Err.Description
    SauvegarderCalendrierCapacite = False
End Function



' -------------------------------------------------------------
' LIAISON DYNAMIQUE DES BOUTONS DE LA FEUILLE ACCUEIL
' -------------------------------------------------------------
Public Sub LierBoutonsAccueil()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    If ws Is Nothing Then Exit Sub
    
    Dim wbN As String
    wbN = "'" & ThisWorkbook.Name & "'!"
    
    ws.Shapes("btn_ImporterDonnees").OnAction = wbN & "ImporterDonnees"
    ws.Shapes("btn_LancerForecast").OnAction = wbN & "LancerForecast"
    ws.Shapes("btn_LancerDemandPlan").OnAction = wbN & "LancerDemandPlan"
    ws.Shapes("btn_LancerSupplyPlan").OnAction = wbN & "LancerSupplyPlan"
    ws.Shapes("btn_GenererScenarios").OnAction = wbN & "GenererScenarios"
    ws.Shapes("btn_CalculerKPI").OnAction = wbN & "CalculerKPI"
    ws.Shapes("btn_ValiderPlanFinal").OnAction = wbN & "ValiderPlanFinal"
    ws.Shapes("btn_LancerMonteCarlo").OnAction = wbN & "LancerMonteCarlo"
    ws.Shapes("btn_LancerCycleComplet").OnAction = wbN & "LancerCycleComplet"
    ws.Shapes("btn_ExporterVersPowerBI").OnAction = wbN & "ExporterVersPowerBI"
    ws.Shapes("btn_ReinitialiserCalculs").OnAction = wbN & "ReinitialiserCalculs"
    ws.Shapes("btn_AfficherLog").OnAction = wbN & "AfficherLog"
    ws.Shapes("btn_Gestion_Produits").OnAction = wbN & "OuvrirGestionProduits"
    ws.Shapes("btn_Commandes_Clients").OnAction = wbN & "OuvrirCommandesClients"
    ws.Shapes("btn_Mouvements_Stocks").OnAction = wbN & "OuvrirMouvementsStock"
    ws.Shapes("btn_Ajustements_Commerciaux").OnAction = wbN & "OuvrirAjustementsCommerciaux"
    ws.Shapes("btn_Parametres_Couts").OnAction = wbN & "OuvrirGestionParametres"
    ws.Shapes("btn_Tableau_Pilotage").OnAction = wbN & "AfficherTableauDeBordPilotage"
    ws.Shapes("btn_Export_PDF").OnAction = wbN & "ExporterRapportPDF"
    On Error GoTo 0
End Sub

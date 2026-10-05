Attribute VB_Name = "mod04_PLANS"
Option Explicit
'==============================================================
' MODULE PLANS - Demand Plan, Capacity Plan, Inventory Plan, Gap Analysis
' F4/F5/F6/F7 - EF-04 a EF-07 - AtlasFood S&OP DSS
'==============================================================

'-------------------------------------------------------------
' DEMAND PLAN (F4) - Agreger le forecast retenu par SKU/mois
'-------------------------------------------------------------
Public Sub LancerDemandPlan()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LogMsg "DEMANDPLAN", "INFO", "=== DEBUT DEMAND PLAN ==="
    ShowProgress "Demand Plan...", 5

    Dim wsFc As Worksheet: Set wsFc = ThisWorkbook.Sheets(SH_FORECAST)
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)

    If wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Lancez d'abord le Forecast (J2).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If

    ClearSheet SH_DEMAND_PLAN
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cycleID As String: cycleID = GetCycleID()
    Dim lastFcRow As Long: lastFcRow = wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastFcRow
        If UCase(CStr(wsFc.Cells(r, 7).Value)) <> "RETENUE" Then GoTo NextRow
        Dim sku As String: sku = CStr(wsFc.Cells(r, 1).Value)
        Dim pd As Date: pd = wsFc.Cells(r, 2).Value
        Dim qty As Double
        If IsNumeric(wsFc.Cells(r, 3).Value) Then qty = CDbl(wsFc.Cells(r, 3).Value)
        Dim key As String: key = sku & "|" & Format(pd, "YYYY-MM")
        If dict.Exists(key) Then dict(key) = dict(key) + qty Else dict.Add key, qty
NextRow:
    Next r

    Dim rowOut As Long: rowOut = 2
    Dim k As Variant
    For Each k In dict.Keys()
        Dim parts() As String: parts = Split(CStr(k), "|")
        wsDp.Cells(rowOut, 1).Value = parts(0)
        wsDp.Cells(rowOut, 2).Value = ParseYM(parts(1))
        wsDp.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsDp.Cells(rowOut, 3).Value = Round(CDbl(dict(k)), 0)
        wsDp.Cells(rowOut, 4).Value = 0
        wsDp.Cells(rowOut, 5).Formula = "=C" & rowOut & "+D" & rowOut
        wsDp.Cells(rowOut, 6).Value = cycleID
        wsDp.Cells(rowOut, 7).Value = "Forecast auto"
        rowOut = rowOut + 1
    Next k

    Dim dpHdrs As Variant
    dpHdrs = Array("SKU_ID", "Mois", "Qte_stat", "Ajustement", "Qte_finale", "Cycle_ID", "Commentaire")
    SetupTable wsDp, "tbl_DEMAND_PLAN", dpHdrs, rowOut, "TableStyleMedium4"

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Demand Plan calcule - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "DEMANDPLAN", "INFO", "OK - " & (rowOut - 2) & " lignes"
    If Application.UserControl Then
        MsgBox "Demand Plan genere (" & (rowOut - 2) & " lignes)." & vbCrLf & _
               "Vous pouvez ajuster la colonne Ajustement (D) si necessaire." & vbCrLf & _
               "Etape suivante : Capacity + Inventory + Gap (J3).", vbInformation, TOOL_NAME
    End If
    wsDp.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "DEMANDPLAN", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Demand Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' CAPACITY PLAN (F5)
'-------------------------------------------------------------
Public Sub LancerCapacityPlan()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LogMsg "CAPACITY", "INFO", "=== DEBUT CAPACITY PLAN ==="
    ShowProgress "Capacity Plan...", 5

    Dim wsDp As Worksheet:    Set wsDp    = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsCap As Worksheet:   Set wsCap   = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsPlant As Worksheet: Set wsPlant = ThisWorkbook.Sheets(SH_PLANT)
    Dim wsProd As Worksheet:  Set wsProd  = ThisWorkbook.Sheets(SH_PRODUCT)

    If wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Lancez d'abord le Demand Plan (J2b).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    ClearSheet SH_CAPACITY

    Dim cycleID As String:    cycleID = GetCycleID()
    Dim seuil As Double:      seuil   = CDbl(GetParam("SEUIL_ALERTE_GAP"))

    ' Capacites par usine
    Dim dictCap As Object: Set dictCap = CreateObject("Scripting.Dictionary")
    Dim loPlant As ListObject
    On Error Resume Next: Set loPlant = wsPlant.ListObjects("tbl_PLANT"): On Error GoTo 0
    If Not loPlant Is Nothing And Not loPlant.DataBodyRange Is Nothing Then
        Dim pr As Long
        For pr = 1 To loPlant.DataBodyRange.Rows.Count
            Dim pid As String:  pid  = CStr(loPlant.DataBodyRange(pr, 1).Value)
            Dim cRef As Double: cRef = CDbl(loPlant.DataBodyRange(pr, 3).Value)
            Dim mg As Double:   mg   = CDbl(loPlant.DataBodyRange(pr, 4).Value)
            Dim hsPct As Double: hsPct = CDbl(loPlant.DataBodyRange(pr, 5).Value)
            Dim stAbs As Double: stAbs = CDbl(loPlant.DataBodyRange(pr, 6).Value)
            If Not dictCap.Exists(pid) Then dictCap.Add pid, (cRef * (1 + mg)) & "|" & (cRef * hsPct) & "|" & stAbs
        Next pr
    End If

    ' Mapping SKU->Plant
    Dim spMap As Object: Set spMap = CreateObject("Scripting.Dictionary")
    Dim loProd As ListObject
    On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
        Dim pp As Long
        For pp = 1 To loProd.DataBodyRange.Rows.Count
            Dim sid As String: sid = CStr(loProd.DataBodyRange(pp, 1).Value)
            Dim pld As String: pld = CStr(loProd.DataBodyRange(pp, 4).Value)
            If Not spMap.Exists(sid) Then spMap.Add sid, pld
        Next pp
    End If

    ' Agregation charge par usine/mois
    Dim dictCharge As Object: Set dictCharge = CreateObject("Scripting.Dictionary")
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim dr As Long
    For dr = 2 To dpLast
        Dim dSku As String: dSku = CStr(wsDp.Cells(dr, 1).Value)
        Dim dDate As Date: dDate = wsDp.Cells(dr, 2).Value
        Dim dQty As Double
        If IsNumeric(wsDp.Cells(dr, 5).Value) Then dQty = CDbl(wsDp.Cells(dr, 5).Value)
        Dim dPlant As String: If spMap.Exists(dSku) Then dPlant = CStr(spMap(dSku)) Else dPlant = "INCONNU"
        Dim ck As String: ck = dPlant & "|" & Format(dDate, "YYYY-MM")
        If dictCharge.Exists(ck) Then dictCharge(ck) = dictCharge(ck) + dQty Else dictCharge.Add ck, dQty
    Next dr

    Dim allKeys As Variant: allKeys = dictCharge.Keys()
    Dim allItems As Variant: allItems = dictCharge.Items()
    Dim rowOut As Long: rowOut = 2
    Dim ik As Long
    For ik = 0 To UBound(allKeys)
        Dim kp() As String: kp = Split(CStr(allKeys(ik)), "|")
        Dim kPlant As String: kPlant = kp(0)
        Dim kMois As Date: kMois = ParseYM(kp(1))
        Dim charge As Double: charge = CDbl(allItems(ik))
        Dim capN As Double: Dim capHS As Double: Dim capST As Double
        If dictCap.Exists(kPlant) Then
            Dim cParts() As String: cParts = Split(CStr(dictCap(kPlant)), "|")
            capN = CDbl(cParts(0)): capHS = CDbl(cParts(1)): capST = CDbl(cParts(2))
        Else
            capN = 5000: capHS = 1000: capST = 500
        End If
        Dim util As Double: If capN > 0 Then util = charge / capN Else util = 0
        Dim st As String
        If util > 1 Then
            st = "SURCHARGE"
        ElseIf util > seuil Then
            st = "ALERTE"
        Else
            st = "OK"
        End If
        wsCap.Cells(rowOut, 1).Value = kPlant
        wsCap.Cells(rowOut, 2).Value = kMois: wsCap.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsCap.Cells(rowOut, 3).Value = Round(capN, 0)
        wsCap.Cells(rowOut, 4).Value = Round(capHS, 0)
        wsCap.Cells(rowOut, 5).Value = Round(capST, 0)
        wsCap.Cells(rowOut, 6).Value = Round(charge, 0)
        wsCap.Cells(rowOut, 7).Value = Round(util * 100, 1)
        wsCap.Cells(rowOut, 8).Value = st
        wsCap.Cells(rowOut, 9).Value = cycleID
        Select Case st
            Case "SURCHARGE": wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(255, 150, 150)
            Case "ALERTE":    wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(255, 230, 150)
            Case Else:        wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        rowOut = rowOut + 1
    Next ik

    Dim capHdrs As Variant
    capHdrs = Array("Plant_ID", "Mois", "Capacite_normale", "Cap_HS_max", "Cap_ST_max", "Charge", "Utilisation_pct", "Statut", "Cycle_ID")
    SetupTable wsCap, "tbl_CAPACITY", capHdrs, rowOut, "TableStyleMedium5"

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Capacity Plan calcule - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "CAPACITY", "INFO", "OK - " & (rowOut - 2) & " lignes"
    If Application.UserControl Then MsgBox "Capacity Plan genere.", vbInformation, TOOL_NAME
    wsCap.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "CAPACITY", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Capacity Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' INVENTORY PLAN (F6)
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
        If Application.UserControl Then MsgBox "Lancez d'abord le Demand Plan (J2b).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    ClearSheet SH_INVENTORY

    Dim cycleID As String: cycleID = GetCycleID()
    Dim nivServ As Double: nivServ = CDbl(GetParam("NIVEAU_SERVICE"))
    Dim delai As Double:   delai   = CDbl(GetParam("DELAI_LIVRAISON"))
    Dim seuil2 As Double:  seuil2  = CDbl(GetParam("SEUIL_SURSTOCK"))

    ' Z-score selon niveau de service
    Dim Z As Double
    Select Case True
        Case nivServ >= 0.99: Z = 2.33
        Case nivServ >= 0.98: Z = 2.05
        Case nivServ >= 0.95: Z = 1.65
        Case nivServ >= 0.90: Z = 1.28
        Case Else: Z = 1.0
    End Select

    ' DLC depuis PRODUCT
    Dim dlcMap As Object: Set dlcMap = CreateObject("Scripting.Dictionary")
    Dim loProd As ListObject
    On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
        Dim pp As Long
        For pp = 1 To loProd.DataBodyRange.Rows.Count
            Dim sID As String: sID = CStr(loProd.DataBodyRange(pp, 1).Value)
            Dim dlcJ As Integer
            If IsNumeric(loProd.DataBodyRange(pp, 6).Value) Then dlcJ = CInt(loProd.DataBodyRange(pp, 6).Value)
            If Not dlcMap.Exists(sID) Then dlcMap.Add sID, dlcJ
        Next pp
    End If

    ' Stocks initiaux
    Dim stockInit As Object: Set stockInit = CreateObject("Scripting.Dictionary")
    stockInit("SKU-01") = CDbl(GetParam("STOCK_INIT_SKU01"))
    stockInit("SKU-02") = CDbl(GetParam("STOCK_INIT_SKU02"))
    stockInit("SKU-03") = CDbl(GetParam("STOCK_INIT_SKU03"))
    stockInit("SKU-04") = CDbl(GetParam("STOCK_INIT_SKU04"))
    stockInit("SKU-05") = CDbl(GetParam("STOCK_INIT_SKU05"))

    ' Collecter et trier les lignes Demand Plan
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim n As Long: n = dpLast - 1
    If n <= 0 Then GoTo CleanExit

    Dim dpKeys() As String: ReDim dpKeys(n - 1)
    Dim dpQty() As Double:  ReDim dpQty(n - 1)
    Dim dr As Long
    For dr = 0 To n - 1
        Dim row As Long: row = dr + 2
        Dim dSku As String: dSku = CStr(wsDp.Cells(row, 1).Value)
        Dim dDate As Date: dDate = wsDp.Cells(row, 2).Value
        Dim dQ As Double
        If IsNumeric(wsDp.Cells(row, 5).Value) Then dQ = CDbl(wsDp.Cells(row, 5).Value)
        dpKeys(dr) = dSku & "|" & Format(dDate, "YYYY-MM")
        dpQty(dr) = dQ
    Next dr

    ' Tri par cle (SKU puis mois)
    Dim i As Long, j As Long, tmpK As String, tmpQ As Double
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If dpKeys(i) > dpKeys(j) Then
                tmpK = dpKeys(i): dpKeys(i) = dpKeys(j): dpKeys(j) = tmpK
                tmpQ = dpQty(i): dpQty(i) = dpQty(j): dpQty(j) = tmpQ
            End If
        Next j
    Next i

    ' Projection stock
    Dim currStock As Object: Set currStock = CreateObject("Scripting.Dictionary")
    Dim rowOut As Long: rowOut = 2
    For i = 0 To n - 1
        Dim kparts() As String: kparts = Split(dpKeys(i), "|")
        Dim iSku As String: iSku = kparts(0)
        Dim iMois As Date: iMois = ParseYM(kparts(1))
        Dim demand As Double: demand = dpQty(i)

        Dim stockDeb As Double
        If currStock.Exists(iSku) Then
            stockDeb = CDbl(currStock(iSku))
        ElseIf stockInit.Exists(iSku) Then
            stockDeb = CDbl(stockInit(iSku))
        Else
            stockDeb = 1000
        End If

        Dim cv As Double: cv = 0.2
        Dim ss As Double: ss = Round(Z * cv * demand * Sqr(delai / 30), 0)
        Dim prod As Double: prod = demand
        Dim livr As Double: livr = demand
        Dim stockFin As Double: stockFin = stockDeb + prod - livr

        Dim dlcRest As Double
        If dlcMap.Exists(iSku) Then dlcRest = CLng(dlcMap(iSku)) - 30

        Dim statInv As String
        If stockFin < 0 Then
            statInv = "RUPTURE"
        ElseIf stockFin < ss Then
            statInv = "SOUS_SEUIL"
        ElseIf demand > 0 And stockFin > demand * seuil2 Then
            statInv = "SURSTOCK"
        Else
            statInv = "OK"
        End If

        wsInv.Cells(rowOut, 1).Value = iSku
        wsInv.Cells(rowOut, 2).Value = iMois: wsInv.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsInv.Cells(rowOut, 3).Value = Round(stockDeb, 0)
        wsInv.Cells(rowOut, 4).Value = Round(ss, 0)
        wsInv.Cells(rowOut, 5).Value = Round(prod, 0)
        wsInv.Cells(rowOut, 6).Value = Round(livr, 0)
        wsInv.Cells(rowOut, 7).Value = Round(stockFin, 0)
        wsInv.Cells(rowOut, 8).Value = Round(dlcRest, 0)
        wsInv.Cells(rowOut, 9).Value = statInv
        wsInv.Cells(rowOut, 10).Value = cycleID
        Select Case statInv
            Case "RUPTURE":    wsInv.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(255, 100, 100)
            Case "SOUS_SEUIL": wsInv.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(255, 200, 150)
            Case "SURSTOCK":   wsInv.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(200, 200, 255)
            Case Else:          wsInv.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        If currStock.Exists(iSku) Then currStock(iSku) = stockFin Else currStock.Add iSku, stockFin
        rowOut = rowOut + 1
    Next i

    Dim invHdrs As Variant
    invHdrs = Array("SKU_ID", "Mois", "Stock_debut", "Stock_secu", "Production", "Livraisons", "Stock_fin", "DLC_restante_est", "Statut", "Cycle_ID")
    SetupTable wsInv, "tbl_INVENTORY", invHdrs, rowOut, "TableStyleMedium6"

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Inventory Plan calcule - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "INVENTORY", "INFO", "OK - " & (rowOut - 2) & " lignes"
    If Application.UserControl Then MsgBox "Inventory Plan genere.", vbInformation, TOOL_NAME
    wsInv.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "INVENTORY", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Inventory Plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

'-------------------------------------------------------------
' GAP ANALYSIS (F7)
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
        If Application.UserControl Then MsgBox "Lancez d'abord le Capacity Plan (J3).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    ClearSheet SH_GAP

    Dim cycleID As String: cycleID = GetCycleID()
    Dim seuil As Double: seuil = CDbl(GetParam("SEUIL_ALERTE_GAP"))
    Dim capLast As Long: capLast = wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row
    Dim rowOut As Long: rowOut = 2
    Dim cr As Long
    For cr = 2 To capLast
        Dim capN As Double:   If IsNumeric(wsCap.Cells(cr, 3).Value) Then capN   = CDbl(wsCap.Cells(cr, 3).Value)
        Dim capHS As Double:  If IsNumeric(wsCap.Cells(cr, 4).Value) Then capHS  = CDbl(wsCap.Cells(cr, 4).Value)
        Dim capST As Double:  If IsNumeric(wsCap.Cells(cr, 5).Value) Then capST  = CDbl(wsCap.Cells(cr, 5).Value)
        Dim charge As Double: If IsNumeric(wsCap.Cells(cr, 6).Value) Then charge = CDbl(wsCap.Cells(cr, 6).Value)
        Dim capD As Double: capD = capN + capHS + capST
        Dim gap As Double: gap = capD - charge
        Dim gapPct As Double: If capD > 0 Then gapPct = gap / capD
        Dim st As String
        If gap < 0 Then
            st = "DEFICIT"
        ElseIf charge / Application.Max(1, capN) > seuil Then
            st = "TENSION"
        Else
            st = "OK"
        End If
        wsGap.Cells(rowOut, 1).Value = wsCap.Cells(cr, 1).Value
        wsGap.Cells(rowOut, 2).Value = wsCap.Cells(cr, 2).Value: wsGap.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsGap.Cells(rowOut, 3).Value = Round(charge, 0)
        wsGap.Cells(rowOut, 4).Value = Round(capD, 0)
        wsGap.Cells(rowOut, 5).Value = Round(gap, 0)
        wsGap.Cells(rowOut, 6).Value = Round(gapPct * 100, 1)
        wsGap.Cells(rowOut, 7).Value = st
        wsGap.Cells(rowOut, 8).Value = cycleID
        Select Case st
            Case "DEFICIT": wsGap.Range("A" & rowOut & ":H" & rowOut).Interior.Color = RGB(255, 100, 100)
            Case "TENSION": wsGap.Range("A" & rowOut & ":H" & rowOut).Interior.Color = RGB(255, 220, 120)
            Case Else:       wsGap.Range("A" & rowOut & ":H" & rowOut).Interior.Color = RGB(200, 240, 200)
        End Select
        rowOut = rowOut + 1
    Next cr

    Dim gapHdrs As Variant
    gapHdrs = Array("Plant_ID", "Mois", "Besoin_net", "Cap_disponible", "Gap", "Gap_pct", "Statut", "Cycle_ID")
    SetupTable wsGap, "tbl_GAP", gapHdrs, rowOut, "TableStyleMedium3"

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Gap Analysis terminee - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "GAP", "INFO", "OK - " & (rowOut - 2) & " lignes"
    If Application.UserControl Then
        MsgBox "Gap Analysis terminee." & vbCrLf & "Etape suivante : Scenarios (J4).", vbInformation, TOOL_NAME
    End If
    wsGap.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "GAP", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Gap Analysis: " & Err.Description, vbCritical, TOOL_NAME
End Sub

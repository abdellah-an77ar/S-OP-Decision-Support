Attribute VB_Name = "mod05_SCENARIOS"
Option Explicit
'==============================================================
' MODULE SCENARIOS - Gestionnaire de 5 scenarios S&OP compares
' F10 / EF-14 - AtlasFood S&OP DSS (niveau Avance)
'==============================================================

Public Sub GenererScenarios()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LogMsg "SCENARIOS", "INFO", "=== DEBUT GENERATION SCENARIOS ==="
    ShowProgress "Generation scenarios...", 5

    Dim wsGap As Worksheet: Set wsGap = ThisWorkbook.Sheets(SH_GAP)
    If wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Lancez d'abord la Gap Analysis (J3).", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If

    Dim wsSc As Worksheet:   Set wsSc   = ThisWorkbook.Sheets(SH_SCENARIOS)
    Dim wsCost As Worksheet: Set wsCost = ThisWorkbook.Sheets(SH_COST)
    Dim wsDp As Worksheet:   Set wsDp   = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    ClearSheet SH_SCENARIOS

    Dim cycleID As String: cycleID = GetCycleID()

    ' Agregats Gap
    Dim totCharge As Double, totCap As Double, totGap As Double
    Dim nDef As Integer
    Dim gapLast As Long: gapLast = wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row
    Dim gr As Long
    For gr = 2 To gapLast
        Dim gC As Double: If IsNumeric(wsGap.Cells(gr, 3).Value) Then gC = CDbl(wsGap.Cells(gr, 3).Value)
        Dim gD As Double: If IsNumeric(wsGap.Cells(gr, 4).Value) Then gD = CDbl(wsGap.Cells(gr, 4).Value)
        Dim gG As Double: If IsNumeric(wsGap.Cells(gr, 5).Value) Then gG = CDbl(wsGap.Cells(gr, 5).Value)
        totCharge = totCharge + gC: totCap = totCap + gD: totGap = totGap + gG
        If CStr(wsGap.Cells(gr, 7).Value) = "DEFICIT" Then nDef = nDef + 1
    Next gr
    Dim defAbs As Double: defAbs = Application.Max(0, -totGap)

    ' Couts unitaires moyens
    Dim cHS As Double:   cHS   = GetCoutMoyen("HS",         wsCost)
    Dim cST As Double:   cST   = GetCoutMoyen("ST",         wsCost)
    Dim cStk As Double:  cStk  = GetCoutMoyen("Stockage",   wsCost)
    Dim cPen As Double:  cPen  = GetCoutMoyen("Penurie",    wsCost)
    Dim cProd As Double: cProd = GetCoutMoyen("Prod",       wsCost)

    ' Demande totale horizon
    Dim demTot As Double
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim dr As Long
    For dr = 2 To dpLast
        If IsNumeric(wsDp.Cells(dr, 5).Value) Then demTot = demTot + CDbl(wsDp.Cells(dr, 5).Value)
    Next dr
    If demTot <= 0 Then demTot = 50000

    ' ========= 5 SCENARIOS =========
    Dim rowOut As Long: rowOut = 2
    Dim sc As Integer

    ' SC-01 BASELINE
    sc = 1
    Dim cost1 As Double: cost1 = demTot * cProd + defAbs * cPen
    Dim svc1 As Double: If totCharge > 0 Then svc1 = Application.Max(0, (totCharge - defAbs) / totCharge) * 100 Else svc1 = 100
    Dim util1 As Double: If totCap > 0 Then util1 = totCharge / totCap * 100 Else util1 = 100
    EcrireScenario wsSc, rowOut, "SC-01", "Baseline (Statu quo)", _
        "Aucun levier. Production a capacite nominale.", _
        cost1, svc1, defAbs * 0.1, defAbs * 0.05, util1, cycleID, "Brouillon"
    wsSc.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(220, 220, 220)
    rowOut = rowOut + 1

    ' SC-02 HEURES SUP +20%
    sc = 2
    Dim addCap2 As Double: addCap2 = totCap * 0.2
    Dim def2 As Double: def2 = Application.Max(0, defAbs - addCap2)
    Dim cost2 As Double: cost2 = demTot * cProd + addCap2 * cHS + def2 * cPen
    Dim svc2 As Double: If totCharge > 0 Then svc2 = Application.Max(0, (totCharge - def2) / totCharge) * 100 Else svc2 = 100
    Dim util2 As Double: If totCap * 1.2 > 0 Then util2 = totCharge / (totCap * 1.2) * 100 Else util2 = 80
    EcrireScenario wsSc, rowOut, "SC-02", "Heures supplementaires +20%", _
        "Activation HS sur usines en deficit. Capacite +20%.", _
        cost2, svc2, def2 * 0.1, addCap2 * 0.03, util2, cycleID, "Brouillon"
    wsSc.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(200, 220, 255)
    rowOut = rowOut + 1

    ' SC-03 SOUS-TRAITANCE 50%
    sc = 3
    Dim addCap3 As Double: addCap3 = defAbs * 0.5
    Dim def3 As Double: def3 = Application.Max(0, defAbs - addCap3)
    Dim cost3 As Double: cost3 = demTot * cProd + addCap3 * cST + def3 * cPen
    Dim svc3 As Double: If totCharge > 0 Then svc3 = Application.Max(0, (totCharge - def3) / totCharge) * 100 Else svc3 = 100
    Dim util3 As Double: util3 = util1  ' Capacite interne inchangee
    EcrireScenario wsSc, rowOut, "SC-03", "Sous-traitance partielle (50%)", _
        "Externalisation 50% du deficit. Delai +5j.", _
        cost3, svc3, def3 * 0.1, addCap3 * 0.03, util3, cycleID, "Brouillon"
    wsSc.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(255, 235, 205)
    rowOut = rowOut + 1

    ' SC-04 HS + ST (solution mixte recommandee)
    sc = 4
    Dim addCap4 As Double: addCap4 = totCap * 0.15 + defAbs * 0.4
    Dim def4 As Double: def4 = Application.Max(0, defAbs - addCap4)
    Dim cost4 As Double: cost4 = demTot * cProd + totCap * 0.15 * cHS + defAbs * 0.4 * cST + def4 * cPen
    Dim svc4 As Double: If totCharge > 0 Then svc4 = Application.Max(0, (totCharge - def4) / totCharge) * 100 Else svc4 = 100
    Dim capTot4 As Double: capTot4 = totCap * 1.15 + defAbs * 0.4
    Dim util4 As Double: If capTot4 > 0 Then util4 = totCharge / capTot4 * 100 Else util4 = 75
    EcrireScenario wsSc, rowOut, "SC-04", "HS 15% + Sous-traitance 40% (RECOMMANDE)", _
        "Mixte HS+ST. Equilibre cout/service optimal.", _
        cost4, svc4, def4 * 0.1, (totCap * 0.15 + defAbs * 0.4) * 0.03, util4, cycleID, "Brouillon"
    wsSc.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(180, 255, 180)
    rowOut = rowOut + 1

    ' SC-05 LISSAGE DEMANDE -10%
    sc = 5
    Dim demLisse As Double: demLisse = demTot * 0.9
    Dim def5 As Double: def5 = Application.Max(0, defAbs - demTot * 0.1)
    Dim cost5 As Double: cost5 = demLisse * cProd + def5 * cPen + (demTot - demLisse) * cStk
    Dim svc5 As Double: If totCharge > 0 Then svc5 = Application.Max(0, (demLisse - def5) / totCharge) * 100 Else svc5 = 90
    Dim util5 As Double: If totCap > 0 Then util5 = demLisse / totCap * 100 Else util5 = 90
    EcrireScenario wsSc, rowOut, "SC-05", "Lissage demande -10% (report promotions)", _
        "Revision a la baisse de 10% de la demande.", _
        cost5, svc5, def5 * 0.1, demTot * 0.1 * 0.02, util5, cycleID, "Brouillon"
    wsSc.Range("A" & rowOut & ":J" & rowOut).Interior.Color = RGB(255, 255, 200)
    rowOut = rowOut + 1

    Dim scHdrs As Variant
    scHdrs = Array("Scenario_ID", "Nom", "Hypotheses", "Cout_total", "Taux_service_pct", "Stock_final_tot", "Risque_peremption", "Util_max_pct", "Cycle_ID", "Statut")
    SetupTable wsSc, "tbl_SCENARIOS", scHdrs, rowOut, "TableStyleMedium7"
    wsSc.Columns("C").ColumnWidth = 50
    wsSc.Columns("B").ColumnWidth = 35

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "5 scenarios generes - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "SCENARIOS", "INFO", "5 scenarios generes."
    If Application.UserControl Then
        MsgBox "5 scenarios compares dans SCENARIOS." & vbCrLf & _
               "Comparez les colonnes Cout, Service, Utilisation." & vbCrLf & _
               "Calculez les KPI puis validez le plan (J5).", vbInformation, TOOL_NAME
    End If
    wsSc.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "SCENARIOS", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Scenarios: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Private Sub EcrireScenario(ws As Worksheet, rowOut As Long, _
    scID As String, nom As String, hypo As String, _
    cout As Double, service As Double, stockFin As Double, _
    risquePer As Double, utilMax As Double, cycleID As String, statut As String)
    ws.Cells(rowOut, 1).Value = scID
    ws.Cells(rowOut, 2).Value = nom
    ws.Cells(rowOut, 3).Value = hypo
    ws.Cells(rowOut, 4).Value = Round(cout, 0)
    ws.Cells(rowOut, 5).Value = Round(service, 1)
    ws.Cells(rowOut, 6).Value = Round(stockFin, 0)
    ws.Cells(rowOut, 7).Value = Round(risquePer, 0)
    ws.Cells(rowOut, 8).Value = Round(utilMax, 1)
    ws.Cells(rowOut, 9).Value = cycleID
    ws.Cells(rowOut, 10).Value = statut
End Sub

Private Function GetCoutMoyen(typeStr As String, wsCost As Worksheet) As Double
    Dim lo As ListObject
    On Error Resume Next: Set lo = wsCost.ListObjects("tbl_COST"): On Error GoTo 0
    Dim total As Double: Dim cnt As Integer: cnt = 0
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.Count
            If CStr(lo.DataBodyRange(r, 1).Value) = typeStr Then
                If IsNumeric(lo.DataBodyRange(r, 3).Value) Then
                    total = total + CDbl(lo.DataBodyRange(r, 3).Value): cnt = cnt + 1
                End If
            End If
        Next r
    End If
    If cnt > 0 Then GetCoutMoyen = total / cnt Else GetCoutMoyen = 300
End Function

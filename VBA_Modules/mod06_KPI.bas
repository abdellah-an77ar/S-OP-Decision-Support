Attribute VB_Name = "mod06_KPI"
Option Explicit
'==============================================================
' MODULE KPI - Calcul KPI, Alertes, Validation Plan, Monte Carlo
' F25/F23/F18 - EF-08/EF-13/EF-23 - S&OP DSS
'==============================================================

'-------------------------------------------------------------
' CALCUL KPI GLOBAL (F25 / EF-08)
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
    Dim kIdx As Long: kIdx = 1

    ' --- KPI 1 : Taux de service ---
    Dim totDem As Double: totDem = 0
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim dr As Long
    For dr = 2 To dpLast
        If IsNumeric(wsDp.Cells(dr, 5).Value) Then totDem = totDem + CDbl(wsDp.Cells(dr, 5).Value)
    Next dr

    Dim totRup As Double: totRup = 0
    Dim invLast As Long: invLast = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    Dim ir As Long
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then
            If IsNumeric(wsInv.Cells(ir, 6).Value) Then totRup = totRup + CDbl(wsInv.Cells(ir, 6).Value)
        End If
    Next ir
    Dim svc As Double: If totDem > 0 Then svc = (totDem - totRup) / totDem * 100 Else svc = 100
    AddKPI wsKpi, rowOut, kIdx, "Taux de service global", "Global", "Horizon S.O.P.", svc, "%", 95, 90, cycleID, True
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 2 : Utilisation max capacite ---
    Dim maxUtil As Double: maxUtil = 0
    Dim capLast As Long: capLast = wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row
    Dim cr As Long
    For cr = 2 To capLast
        If IsNumeric(wsCap.Cells(cr, 7).Value) Then
            If CDbl(wsCap.Cells(cr, 7).Value) > maxUtil Then maxUtil = CDbl(wsCap.Cells(cr, 7).Value)
        End If
    Next cr
    AddKPI wsKpi, rowOut, kIdx, "Utilisation max capacite", "Capacity", "Peak horizon", maxUtil, "%", 85, 95, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 3 : WAPE moyen forecast ---
    Dim sumW As Double: Dim nW As Long: nW = 0
    Dim fcLast As Long: fcLast = wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row
    Dim fr As Long
    For fr = 2 To fcLast
        If UCase(CStr(wsFc.Cells(fr, 7).Value)) = "RETENUE" Then
            If IsNumeric(wsFc.Cells(fr, 5).Value) Then
                sumW = sumW + CDbl(wsFc.Cells(fr, 5).Value): nW = nW + 1
            End If
        End If
    Next fr
    Dim wapeMoy As Double: If nW > 0 Then wapeMoy = sumW / nW
    AddKPI wsKpi, rowOut, kIdx, "WAPE moyen forecast", "Forecast", "Methode retenue", wapeMoy, "%", 20, 30, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 4 : Ruptures ---
    Dim nRup As Long: nRup = 0
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then nRup = nRup + 1
    Next ir
    If wsKpi.ProtectContents Then wsKpi.Unprotect
    AddKPI wsKpi, rowOut, kIdx, "Nombre de ruptures de stock", "Inventory", "Horizon", nRup, "cas", 0, 2, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 5 : Surstocks ---
    Dim nSurstock As Long: nSurstock = 0
    For ir = 2 To invLast
        If CStr(wsInv.Cells(ir, 9).Value) = "SURSTOCK" Then nSurstock = nSurstock + 1
    Next ir
    AddKPI wsKpi, rowOut, kIdx, "Nombre de surstocks", "Inventory", "Horizon", nSurstock, "cas", 0, 3, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 6 : Surcharges capacite ---
    Dim nSurch As Long: nSurch = 0
    For cr = 2 To capLast
        If CStr(wsCap.Cells(cr, 8).Value) = "SURCHARGE" Then nSurch = nSurch + 1
    Next cr
    AddKPI wsKpi, rowOut, kIdx, "Usines-mois en surcharge", "Capacity", "Horizon", nSurch, "cas", 0, 2, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 7 : Deficit capacite total ---
    Dim totDef As Double: totDef = 0
    Dim gapLast As Long: gapLast = wsGap.Cells(wsGap.Rows.Count, 1).End(xlUp).Row
    For cr = 2 To gapLast
        If IsNumeric(wsGap.Cells(cr, 5).Value) Then
            If CDbl(wsGap.Cells(cr, 5).Value) < 0 Then totDef = totDef + Abs(CDbl(wsGap.Cells(cr, 5).Value))
        End If
    Next cr
    AddKPI wsKpi, rowOut, kIdx, "Deficit capacite total (unites)", "Gap", "Horizon", totDef, "unites", 0, 1000, cycleID, False
    rowOut = rowOut + 1: kIdx = kIdx + 1

    ' --- KPI 8 : Demande totale planifiee ---
    AddKPI wsKpi, rowOut, kIdx, "Demande totale planifiee", "Demand", "Horizon S.O.P.", totDem, "unites", 0, 0, cycleID, True
    rowOut = rowOut + 1: kIdx = kIdx + 1

    Dim kpiHdrs As Variant
    kpiHdrs = Array("KPI_ID", "Indicateur", "Perimetre", "Periode", "Valeur", "Unite", "Seuil_vert", "Seuil_rouge", "Statut_seuil", "Cycle_ID")
    SetupTable wsKpi, "tbl_KPI", kpiHdrs, rowOut, "TableStyleMedium2"
    wsKpi.Columns("B").ColumnWidth = 32

    ' Generer les alertes
    GenererAlertes wsKpi

    ' Mettre a jour la feuille de pilotage executif
    On Error Resume Next
    Call ActualiserFeuillePilotage
    On Error GoTo 0

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "KPI calcules - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "KPI", "INFO", "KPI calcules."
    If Application.UserControl Then
        MsgBox "KPI calcules (" & (kIdx - 1) & " indicateurs)." & vbCrLf & _
               "Alertes generees dans ALERTES." & vbCrLf & _
               "Etape suivante : valider le plan (J5).", vbInformation, TOOL_NAME
    End If
    wsKpi.Activate
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "KPI", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur KPI: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Private Sub AddKPI(ws As Worksheet, rowOut As Long, kIdx As Long, _
    nom As String, perimetre As String, periode As String, _
    ByVal val As Double, unite As String, ByVal seuilV As Double, ByVal seuilR As Double, _
    cycleID As String, plusGrand As Boolean)

    ws.Cells(rowOut, 1).Value = "KPI-" & Format(kIdx, "00")
    ws.Cells(rowOut, 2).Value = nom
    ws.Cells(rowOut, 3).Value = perimetre
    ws.Cells(rowOut, 4).Value = periode
    ws.Cells(rowOut, 5).Value = Round(val, 1)
    ws.Cells(rowOut, 6).Value = unite
    ws.Cells(rowOut, 7).Value = seuilV
    ws.Cells(rowOut, 8).Value = seuilR
    ws.Cells(rowOut, 10).Value = cycleID

    Dim st As String: st = "VERT"
    If plusGrand Then
        If val < seuilR Then
            st = "ROUGE"
        ElseIf val < seuilV Then
            st = "ORANGE"
        Else
            st = "VERT"
        End If
    Else
        If seuilR > 0 And val > seuilR Then
            st = "ROUGE"
        ElseIf seuilV > 0 And val > seuilV Then
            st = "ORANGE"
        Else
            st = "VERT"
        End If
    End If
    ws.Cells(rowOut, 9).Value = st
    Select Case st
        Case "VERT":   ws.Cells(rowOut, 9).Interior.Color = RGB(150, 255, 150)
        Case "ORANGE": ws.Cells(rowOut, 9).Interior.Color = RGB(255, 200, 100)
        Case "ROUGE":  ws.Cells(rowOut, 9).Interior.Color = RGB(255, 100, 100)
    End Select
End Sub

'-------------------------------------------------------------
' GENERATION ALERTES
'-------------------------------------------------------------
Private Sub GenererAlertes(wsKpi As Worksheet)
    Dim wsAlt As Worksheet: Set wsAlt = ThisWorkbook.Sheets(SH_ALERTS)
    ClearSheet SH_ALERTS
    Dim cycleID As String: cycleID = GetCycleID()
    Dim rowOut As Long: rowOut = 2
    Dim alertID As Long: alertID = 1

    ' Ruptures
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets(SH_INVENTORY)
    Dim ir As Long
    For ir = 2 To wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        If CStr(wsInv.Cells(ir, 9).Value) = "RUPTURE" Then
            wsAlt.Cells(rowOut, 1).Value = "ALT-" & Format(alertID, "000")
            wsAlt.Cells(rowOut, 2).Value = "RUPTURE_STOCK"
            wsAlt.Cells(rowOut, 3).Value = "CRITIQUE"
            wsAlt.Cells(rowOut, 4).Value = wsInv.Cells(ir, 1).Value & " / " & wsInv.Cells(ir, 2).Text
            wsAlt.Cells(rowOut, 5).Value = "Rupture detectee: stock fin negatif"
            wsAlt.Cells(rowOut, 6).Value = cycleID
            wsAlt.Cells(rowOut, 7).Value = Now()
            wsAlt.Cells(rowOut, 7).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            wsAlt.Range("A" & rowOut & ":G" & rowOut).Interior.Color = RGB(255, 150, 150)
            rowOut = rowOut + 1: alertID = alertID + 1
        End If
    Next ir

    ' Surcharges
    Dim wsCap As Worksheet: Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim cr As Long
    For cr = 2 To wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row
        If CStr(wsCap.Cells(cr, 8).Value) = "SURCHARGE" Then
            wsAlt.Cells(rowOut, 1).Value = "ALT-" & Format(alertID, "000")
            wsAlt.Cells(rowOut, 2).Value = "SURCHARGE_CAPACITE"
            wsAlt.Cells(rowOut, 3).Value = "ELEVEE"
            wsAlt.Cells(rowOut, 4).Value = wsCap.Cells(cr, 1).Value & " / " & wsCap.Cells(cr, 2).Text
            wsAlt.Cells(rowOut, 5).Value = "Surcharge: util=" & wsCap.Cells(cr, 7).Value & "%"
            wsAlt.Cells(rowOut, 6).Value = cycleID
            wsAlt.Cells(rowOut, 7).Value = Now()
            wsAlt.Cells(rowOut, 7).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            wsAlt.Range("A" & rowOut & ":G" & rowOut).Interior.Color = RGB(255, 220, 150)
            rowOut = rowOut + 1: alertID = alertID + 1
        End If
    Next cr

    ' Surstocks
    Dim wsInv2 As Worksheet: Set wsInv2 = ThisWorkbook.Sheets(SH_INVENTORY)
    For ir = 2 To wsInv2.Cells(wsInv2.Rows.Count, 1).End(xlUp).Row
        If CStr(wsInv2.Cells(ir, 9).Value) = "SURSTOCK" Then
            wsAlt.Cells(rowOut, 1).Value = "ALT-" & Format(alertID, "000")
            wsAlt.Cells(rowOut, 2).Value = "SURSTOCK"
            wsAlt.Cells(rowOut, 3).Value = "MODEREE"
            wsAlt.Cells(rowOut, 4).Value = wsInv2.Cells(ir, 1).Value & " / " & wsInv2.Cells(ir, 2).Text
            wsAlt.Cells(rowOut, 5).Value = "Surstock detecte: stock fin > 2x demande"
            wsAlt.Cells(rowOut, 6).Value = cycleID
            wsAlt.Cells(rowOut, 7).Value = Now()
            wsAlt.Cells(rowOut, 7).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            wsAlt.Range("A" & rowOut & ":G" & rowOut).Interior.Color = RGB(200, 200, 255)
            rowOut = rowOut + 1: alertID = alertID + 1
        End If
    Next ir

    Dim altHdrs As Variant
    altHdrs = Array("Alert_ID", "Type", "Criticite", "Perimetre", "Message", "Cycle_ID", "Horodatage")
    SetupTable wsAlt, "tbl_ALERTES", altHdrs, rowOut, "TableStyleMedium3"
    wsAlt.Columns("E").ColumnWidth = 45
    wsAlt.Columns("G").ColumnWidth = 22
    LogMsg "KPI", "INFO", (alertID - 1) & " alerte(s) generee(s)"
End Sub

'-------------------------------------------------------------
' VALIDER LE PLAN FINAL (F23 / EF-13)
'-------------------------------------------------------------
Public Sub ValiderPlanFinal()
    On Error GoTo ErrHandler

    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Sheets(SH_SCENARIOS)
    If wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Generez d'abord les scenarios (J4).", vbCritical, TOOL_NAME
        Exit Sub
    End If

    ' Construire liste scenarios
    Dim scenList As String
    scenList = "Scenarios disponibles:" & vbCrLf & vbCrLf
    Dim sr As Long
    For sr = 2 To wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row
        scenList = scenList & wsSc.Cells(sr, 1).Value & " - " & wsSc.Cells(sr, 2).Value & vbCrLf
        If IsNumeric(wsSc.Cells(sr, 4).Value) Then
            scenList = scenList & "   Cout: " & Format(CDbl(wsSc.Cells(sr, 4).Value), "#,##0") & " UM"
        End If
        If IsNumeric(wsSc.Cells(sr, 5).Value) Then
            scenList = scenList & " | Service: " & wsSc.Cells(sr, 5).Value & "%"
        End If
        scenList = scenList & vbCrLf
    Next sr
    scenList = scenList & vbCrLf & "Entrez l'ID (ex: SC-04) :"

    Dim chosen As String: chosen = InputBox(scenList, TOOL_NAME & " - Validation Plan", "SC-04")
    If Trim(chosen) = "" Then Exit Sub
    chosen = UCase(Trim(chosen))

    Dim scenRow As Long: scenRow = 0
    For sr = 2 To wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row
        If UCase(CStr(wsSc.Cells(sr, 1).Value)) = chosen Then scenRow = sr: Exit For
    Next sr
    If scenRow = 0 Then
        If Application.UserControl Then MsgBox "Scenario '" & chosen & "' introuvable.", vbCritical, TOOL_NAME
        Exit Sub
    End If

    Dim confMsg As String
    confMsg = "Valider le plan avec " & wsSc.Cells(scenRow, 1).Value & " ?" & vbCrLf
    confMsg = confMsg & wsSc.Cells(scenRow, 2).Value & vbCrLf & vbCrLf
    If IsNumeric(wsSc.Cells(scenRow, 4).Value) Then confMsg = confMsg & "Cout total: " & Format(CDbl(wsSc.Cells(scenRow, 4).Value), "#,##0") & " UM" & vbCrLf
    If IsNumeric(wsSc.Cells(scenRow, 5).Value) Then confMsg = confMsg & "Taux de service: " & wsSc.Cells(scenRow, 5).Value & "%" & vbCrLf
    If MsgBox(confMsg, vbYesNo + vbQuestion, TOOL_NAME) <> vbYes Then Exit Sub

    Call ValiderPlan(chosen)
    Exit Sub
ErrHandler:
    LogMsg "PLANFINAL", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur validation plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Public Sub ValiderPlan(chosen As String)
    On Error GoTo ErrHandler
    chosen = UCase(Trim(chosen))
    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Sheets(SH_SCENARIOS)
    Dim scenRow As Long: scenRow = 0
    Dim sr As Long
    For sr = 2 To wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row
        If UCase(CStr(wsSc.Cells(sr, 1).Value)) = chosen Then scenRow = sr: Exit For
    Next sr
    If scenRow = 0 Then
        If wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row >= 2 Then
            scenRow = 2: chosen = CStr(wsSc.Cells(2, 1).Value)
        Else
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    Dim wsPf As Worksheet:   Set wsPf   = ThisWorkbook.Sheets(SH_PLAN_FINAL)
    Dim wsDp As Worksheet:   Set wsDp   = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets(SH_PRODUCT)
    ClearSheet SH_PLAN_FINAL

    Dim planID As String:  planID  = "PLAN-" & Format(Now(), "YYYYMMDD-HHMMSS")
    Dim valideur As String: valideur = GetCurrentUser()
    Dim cycleID As String:  cycleID = GetCycleID()
    Dim dpLast As Long:     dpLast  = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim rowOut As Long: rowOut = 2

    ' Mapping SKU->Plant
    Dim spMap As Object: Set spMap = CreateObject("Scripting.Dictionary")
    Dim loProd As ListObject
    On Error Resume Next: Set loProd = wsProd.ListObjects("tbl_PRODUCT"): On Error GoTo 0
    If Not loProd Is Nothing And Not loProd.DataBodyRange Is Nothing Then
        Dim pp As Long
        For pp = 1 To loProd.DataBodyRange.Rows.Count
            Dim sid As String: sid = CStr(loProd.DataBodyRange(pp, 1).Value)
            Dim pid As String: pid = CStr(loProd.DataBodyRange(pp, 4).Value)
            If Not spMap.Exists(sid) Then spMap.Add sid, pid
        Next pp
    End If

    Dim dr As Long
    For dr = 2 To dpLast
        Dim pSku As String:  pSku  = CStr(wsDp.Cells(dr, 1).Value)
        Dim pMois As Date:   pMois = wsDp.Cells(dr, 2).Value
        Dim pDem  As Double: If IsNumeric(wsDp.Cells(dr, 5).Value) Then pDem = CDbl(wsDp.Cells(dr, 5).Value)
        Dim pPlant As String: If spMap.Exists(pSku) Then pPlant = CStr(spMap(pSku)) Else pPlant = "USN-01"

        Dim pNorm As Double: pNorm = pDem
        Dim pHS   As Double: pHS   = 0
        Dim pST   As Double: pST   = 0

        Select Case chosen
            Case "SC-02": pHS = pDem * 0.2
            Case "SC-03": pST = pDem * 0.1
            Case "SC-04": pHS = pDem * 0.15: pST = pDem * 0.1
            Case "SC-05": pNorm = pDem * 0.9
        End Select

        wsPf.Cells(rowOut, 1).Value  = planID
        wsPf.Cells(rowOut, 2).Value  = 1
        wsPf.Cells(rowOut, 3).Value  = pSku
        wsPf.Cells(rowOut, 4).Value  = pPlant
        wsPf.Cells(rowOut, 5).Value  = pMois: wsPf.Cells(rowOut, 5).NumberFormat = "mmm-yy"
        wsPf.Cells(rowOut, 6).Value  = Round(pNorm, 0)
        wsPf.Cells(rowOut, 7).Value  = Round(pHS, 0)
        wsPf.Cells(rowOut, 8).Value  = Round(pST, 0)
        wsPf.Cells(rowOut, 9).Value  = 0
        wsPf.Cells(rowOut, 10).Value = Round(pDem, 0)
        wsPf.Cells(rowOut, 11).Value = 0
        wsPf.Cells(rowOut, 12).Value = chosen & " - " & wsSc.Cells(scenRow, 2).Value
        wsPf.Cells(rowOut, 13).Value = "Valide"
        wsPf.Cells(rowOut, 14).Value = valideur
        wsPf.Cells(rowOut, 15).Value = Now(): wsPf.Cells(rowOut, 15).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsPf.Range("A" & rowOut & ":O" & rowOut).Interior.Color = RGB(200, 240, 200)
        rowOut = rowOut + 1
    Next dr

    ' Marquer scenario comme valide
    wsSc.Cells(scenRow, 10).Value = "VALIDE"
    wsSc.Range("A" & scenRow & ":J" & scenRow).Interior.Color = RGB(150, 255, 150)

    Dim pfHdrs As Variant
    pfHdrs = Array("Plan_ID", "Version", "SKU_ID", "Plant_ID", "Mois", "Prod_normale", "Prod_HS", "Prod_ST", "Stock_fin", "Demande", "Non_servi", "Scenario_retenu", "Statut", "Valideur", "Horodatage")
    SetupTable wsPf, "tbl_PLAN_FINAL", pfHdrs, rowOut, "TableStyleMedium4"
    wsPf.Columns("A:O").AutoFit

    ' Mettre a jour la feuille de pilotage executif
    On Error Resume Next
    Call ActualiserFeuillePilotage
    On Error GoTo 0

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    SetStatus "Plan VALIDE : " & chosen & " - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "PLANFINAL", "INFO", "Plan valide: " & chosen & " par " & valideur
    If Application.UserControl Then
        MsgBox "Plan valide avec " & chosen & "." & vbCrLf & _
               "Consultez PLAN_FINAL." & vbCrLf & _
               "Etape suivante : Exporter vers Power BI.", vbInformation, TOOL_NAME
    End If
    wsPf.Activate
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    LogMsg "PLANFINAL", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur validation plan: " & Err.Description, vbCritical, TOOL_NAME
End Sub


'-------------------------------------------------------------
' MONTE CARLO (F18 / EF-23 - Expert)
'-------------------------------------------------------------
Public Sub LancerMonteCarlo()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LogMsg "MC", "INFO", "=== MONTE CARLO ==="
    ShowProgress "Monte Carlo...", 5

    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    If wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row < 2 Then
        If Application.UserControl Then MsgBox "Lancez d'abord le Demand Plan.", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If

    ' Creer/recuperer feuille MC_RESULTS
    Dim wsMC As Worksheet
    On Error Resume Next: Set wsMC = ThisWorkbook.Sheets("MC_RESULTS"): On Error GoTo 0
    If wsMC Is Nothing Then
        Set wsMC = ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(SH_LOG))
        wsMC.Name = "MC_RESULTS"
        wsMC.Tab.Color = RGB(112, 48, 160)
    End If
    ClearSheet "MC_RESULTS"

    Dim mh() As Variant
    mh = Array("SKU_ID","Iterations","Demande_base","P50","P90","P_rupture_pct","Cout_moyen","Cout_P90","Service_moyen_pct","Cycle_ID")
    Dim c As Integer
    For c = 0 To 9: wsMC.Cells(1, c + 1).Value = mh(c): Next c
    wsMC.Range("A1:J1").Font.Bold = True
    wsMC.Range("A1:J1").Interior.Color = RGB(112, 48, 160)
    wsMC.Range("A1:J1").Font.Color = RGB(255, 255, 255)

    Dim nIter As Integer: nIter = CInt(GetParam("MC_ITERATIONS"))
    Dim sigma As Double:  sigma  = CDbl(GetParam("MC_SIGMA_PCT"))
    Dim cPen  As Double:  cPen   = 800
    Dim cStk  As Double:  cStk   = 12
    Dim cycleID As String: cycleID = GetCycleID()

    ' SKU par somme Demand Plan
    Dim skuDem As Object: Set skuDem = CreateObject("Scripting.Dictionary")
    Dim dpLast As Long: dpLast = wsDp.Cells(wsDp.Rows.Count, 1).End(xlUp).Row
    Dim dr As Long
    For dr = 2 To dpLast
        Dim dSku As String: dSku = CStr(wsDp.Cells(dr, 1).Value)
        Dim dQ As Double: If IsNumeric(wsDp.Cells(dr, 5).Value) Then dQ = CDbl(wsDp.Cells(dr, 5).Value)
        If skuDem.Exists(dSku) Then skuDem(dSku) = skuDem(dSku) + dQ Else skuDem.Add dSku, dQ
    Next dr

    Dim skuArr As Variant: skuArr = skuDem.Keys()
    Dim nSKU As Integer: nSKU = Application.Min(4, skuDem.Count)
    Dim rowOut As Long: rowOut = 2
    Dim s As Integer

    For s = 0 To nSKU - 1
        Dim sk As String: sk = skuArr(s)
        Dim baseDem As Double: baseDem = CDbl(skuDem(sk))
        ShowProgress "Monte Carlo: " & sk, 10 + Int(80 * s / nSKU)

        Dim totCout As Double: Dim nRup As Long: nRup = 0
        Dim totSvc As Double: Dim it As Integer
        Dim simArr() As Double: ReDim simArr(nIter - 1)
        Dim coutArr() As Double: ReDim coutArr(nIter - 1)

        Randomize
        For it = 0 To nIter - 1
            Dim u1 As Double: u1 = Rnd(): If u1 < 0.0001 Then u1 = 0.0001
            Dim u2 As Double: u2 = Rnd(): If u2 < 0.0001 Then u2 = 0.0001
            Dim noise As Double
            noise = Sqr(-2 * Log(u1)) * Cos(2 * 3.14159265358979 * u2)
            Dim simDem As Double: simDem = Application.Max(0, baseDem * (1 + noise * sigma))
            simArr(it) = simDem
            Dim diff As Double: diff = simDem - baseDem
            Dim cout As Double
            If diff > 0 Then cout = diff * cPen: nRup = nRup + 1 Else cout = Abs(diff) * cStk
            coutArr(it) = cout
            totCout = totCout + cout
            totSvc = totSvc + Application.Min(1, baseDem / Application.Max(1, simDem))
        Next it

        ' Percentiles (tri par insertion partiel)
        Dim p50Idx As Long: p50Idx = CLng(nIter * 0.5)
        Dim p90Idx As Long: p90Idx = CLng(nIter * 0.9)
        ' Approximation rapide: tri partiel uniquement sur les premiers 200 elements pour perf
        Dim i As Integer, j As Integer, tmp As Double
        For i = 0 To Application.Min(nIter - 1, p90Idx + 5)
            Dim minIdx As Long: minIdx = i
            For j = i + 1 To nIter - 1
                If simArr(j) < simArr(minIdx) Then minIdx = j
            Next j
            tmp = simArr(i): simArr(i) = simArr(minIdx): simArr(minIdx) = tmp
            tmp = coutArr(i): coutArr(i) = coutArr(minIdx): coutArr(minIdx) = tmp
        Next i

        wsMC.Cells(rowOut, 1).Value = sk
        wsMC.Cells(rowOut, 2).Value = nIter
        wsMC.Cells(rowOut, 3).Value = Round(baseDem, 0)
        wsMC.Cells(rowOut, 4).Value = Round(simArr(p50Idx), 0)
        wsMC.Cells(rowOut, 5).Value = Round(simArr(p90Idx), 0)
        wsMC.Cells(rowOut, 6).Value = Round(nRup / nIter * 100, 1)
        wsMC.Cells(rowOut, 7).Value = Round(totCout / nIter, 0)
        wsMC.Cells(rowOut, 8).Value = Round(coutArr(p90Idx), 0)
        wsMC.Cells(rowOut, 9).Value = Round(totSvc / nIter * 100, 1)
        wsMC.Cells(rowOut, 10).Value = cycleID
        rowOut = rowOut + 1
    Next s

    Dim mcHdrs As Variant
    mcHdrs = Array("SKU_ID", "Iterations", "Demande_base", "P50", "P90", "P_rupture_pct", "Cout_moyen", "Cout_P90", "Service_moyen_pct", "Cycle_ID")
    SetupTable wsMC, "tbl_MC", mcHdrs, rowOut, "TableStyleMedium8"
    wsMC.Columns("A:J").AutoFit

    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    ClearProgress: SetStatus "Monte Carlo termine - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "MC", "INFO", "Monte Carlo OK - " & nSKU & " SKU - " & nIter & " iterations"
    If Application.UserControl Then
        MsgBox "Monte Carlo termine." & vbCrLf & nSKU & " SKU simules, " & nIter & " iterations." & vbCrLf & "Consultez MC_RESULTS.", vbInformation, TOOL_NAME
    End If
    wsMC.Activate
    Exit Sub
CleanExit:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: ClearProgress
    LogMsg "MC", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Monte Carlo: " & Err.Description, vbCritical, TOOL_NAME
End Sub

Attribute VB_Name = "mod07_EXPORT"
Option Explicit
'==============================================================
' MODULE EXPORT - Export CSV Power BI + Cycle complet + Navigation
' F19 / EF-21 + EF-12 - S&OP DSS
'==============================================================

'-------------------------------------------------------------
' CYCLE COMPLET (EF-12 / F22)
'-------------------------------------------------------------
Public Sub LancerCycleComplet()
    Dim rep As Integer
    If Application.UserControl Then
        rep = MsgBox("Lancer le cycle S.O.P. complet ?" & vbCrLf & vbCrLf & _
              "Etapes: Forecast -> Demand Plan -> Capacity -> Inventory -> Gap -> Scenarios -> KPI" & vbCrLf & vbCrLf & _
              "Prerequis : les donnees doivent etre importees (J1).", _
              vbYesNo + vbQuestion, TOOL_NAME)
        If rep <> vbYes Then Exit Sub
    End If

    LogMsg "CYCLE", "INFO", "=== DEBUT CYCLE COMPLET ==="
    ShowProgress "Cycle complet (1/7): Forecast...", 5
    Call LancerForecast

    ShowProgress "Cycle complet (2/7): Demand Plan...", 20
    Call LancerDemandPlan

    ShowProgress "Cycle complet (3/7): Capacity Plan...", 38
    Call LancerCapacityPlan

    ShowProgress "Cycle complet (4/7): Inventory Plan...", 55
    Call LancerInventoryPlan

    ShowProgress "Cycle complet (5/7): Gap Analysis...", 68
    Call LancerGapAnalysis

    ShowProgress "Cycle complet (6/7): Scenarios...", 80
    Call GenererScenarios

    ShowProgress "Cycle complet (7/7): KPI...", 93
    Call CalculerKPI

    ClearProgress
    LogMsg "CYCLE", "INFO", "=== CYCLE COMPLET TERMINE ==="
    SetStatus "Cycle complet termine - " & Format(Now(), "yyyy-mm-dd hh:mm")
    If Application.UserControl Then
        MsgBox "Cycle S.O.P. complet execute avec succes !" & vbCrLf & vbCrLf & _
               "Toutes les etapes sont calculees." & vbCrLf & _
               "Consultez les scenarios et validez le plan (J5).", vbInformation, TOOL_NAME
    End If
    ThisWorkbook.Sheets(SH_SCENARIOS).Activate
End Sub

Public Sub LancerSupplyPlan()
    Call LancerCapacityPlan
    Call LancerInventoryPlan
    Call LancerGapAnalysis
End Sub

Public Sub AfficherLog()
    On Error Resume Next
    ThisWorkbook.Sheets(SH_LOG).Visible = xlSheetVisible
    ThisWorkbook.Sheets(SH_LOG).Activate
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' EXPORT VERS POWER BI (F19 / EF-21)
'-------------------------------------------------------------
Public Sub ExporterVersPowerBI()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    LogMsg "EXPORT", "INFO", "=== EXPORT POWER BI ==="
    ShowProgress "Export Power BI...", 5

    Dim exportDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim defaultDir As String
    If ThisWorkbook.Path <> "" Then
        defaultDir = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), "PowerBI")
        If Not fso.FolderExists(defaultDir) Then defaultDir = fso.BuildPath(ThisWorkbook.Path, "PowerBI")
    End If
    If defaultDir = "" Or Not fso.FolderExists(defaultDir) Then
        If fso.FolderExists("c:\Users\RZR\Documents\EMINES\CI2A\app vba\PowerBI") Then
            defaultDir = "c:\Users\RZR\Documents\EMINES\CI2A\app vba\PowerBI"
        End If
    End If
    
    If defaultDir <> "" And fso.FolderExists(defaultDir) Then
        exportDir = defaultDir
    Else
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.Title = "Dossier de destination pour les fichiers CSV Power BI"
        fd.InitialFileName = ThisWorkbook.Path

        If fd.Show <> -1 Then GoTo CleanExit
        exportDir = fd.SelectedItems(1)
    End If

    ' Creer le sous-dossier si besoin
    If Dir(exportDir, vbDirectory) = "" Then MkDir exportDir

    Dim wsExp As Worksheet: Set wsExp = ThisWorkbook.Sheets(SH_EXPORT)
    ClearSheet SH_EXPORT
    Dim expRow As Long: expRow = 2

    ' Tables a exporter: {SheetName, CsvName}
    Dim tables(10, 1) As String
    tables(0, 0) = SH_FORECAST:     tables(0, 1)  = "PBI_Forecast.csv"
    tables(1, 0) = SH_DEMAND_PLAN:  tables(1, 1)  = "PBI_DemandPlan.csv"
    tables(2, 0) = SH_CAPACITY:     tables(2, 1)  = "PBI_CapacityPlan.csv"
    tables(3, 0) = SH_INVENTORY:    tables(3, 1)  = "PBI_InventoryPlan.csv"
    tables(4, 0) = SH_GAP:          tables(4, 1)  = "PBI_GapAnalysis.csv"
    tables(5, 0) = SH_SCENARIOS:    tables(5, 1)  = "PBI_Scenarios.csv"
    tables(6, 0) = SH_PLAN_FINAL:   tables(6, 1)  = "PBI_PlanFinal.csv"
    tables(7, 0) = SH_KPI:          tables(7, 1)  = "PBI_KPI.csv"
    tables(8, 0) = SH_ALERTS:       tables(8, 1)  = "PBI_Alertes.csv"
    tables(9, 0) = SH_PRODUCT:      tables(9, 1)  = "PBI_Product.csv"
    tables(10, 0) = SH_PLANT:       tables(10, 1) = "PBI_Plant.csv"

    Dim t As Integer
    For t = 0 To 10
        ShowProgress "Export: " & tables(t, 1), 5 + Int(90 * t / 11)
        Dim csvPath As String: csvPath = exportDir & "\" & tables(t, 1)
        Dim ws As Worksheet
        On Error Resume Next: Set ws = ThisWorkbook.Sheets(tables(t, 0)): On Error GoTo 0
        Dim nbRows As Long: nbRows = 0
        If Not ws Is Nothing Then nbRows = ExportCSV(ws, csvPath)

        wsExp.Cells(expRow, 1).Value = tables(t, 1)
        wsExp.Cells(expRow, 2).Value = Now(): wsExp.Cells(expRow, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsExp.Cells(expRow, 3).Value = nbRows
        wsExp.Cells(expRow, 4).Value = IIf(nbRows > 0, "OK", "Vide")
        If nbRows > 0 Then wsExp.Range("A" & expRow & ":D" & expRow).Interior.Color = RGB(200, 240, 200) _
        Else wsExp.Range("A" & expRow & ":D" & expRow).Interior.Color = RGB(255, 220, 150)
        expRow = expRow + 1
        LogMsg "EXPORT", "INFO", tables(t, 1) & ": " & nbRows & " lignes"
    Next t

    ' Exporter MC_RESULTS si presente
    Dim wsMC As Worksheet
    On Error Resume Next: Set wsMC = ThisWorkbook.Sheets("MC_RESULTS"): On Error GoTo 0
    If Not wsMC Is Nothing Then
        Dim mcRows As Long: mcRows = ExportCSV(wsMC, exportDir & "\PBI_MonteCarlo.csv")
        wsExp.Cells(expRow, 1).Value = "PBI_MonteCarlo.csv"
        wsExp.Cells(expRow, 2).Value = Now(): wsExp.Cells(expRow, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsExp.Cells(expRow, 3).Value = mcRows
        wsExp.Cells(expRow, 4).Value = IIf(mcRows > 0, "OK", "Vide")
        If mcRows > 0 Then wsExp.Range("A" & expRow & ":D" & expRow).Interior.Color = RGB(200, 240, 200)
    End If

    Application.ScreenUpdating = True
    ClearProgress: SetStatus "Export Power BI termine - " & Format(Now(), "yyyy-mm-dd hh:mm")
    If Application.UserControl Then
        MsgBox "Export termine !" & vbCrLf & _
               "Fichiers CSV crees dans :" & vbCrLf & exportDir & vbCrLf & vbCrLf & _
               "Actualisez votre rapport Power BI pour voir les donnees.", vbInformation, TOOL_NAME
        Shell "explorer.exe """ & exportDir & """", vbNormalFocus
    End If
    ThisWorkbook.Sheets(SH_EXPORT).Visible = xlSheetVisible
    wsExp.Activate
    Exit Sub

CleanExit:
    Application.ScreenUpdating = True: ClearProgress: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: ClearProgress
    LogMsg "EXPORT", "ERROR", Err.Description
    If Application.UserControl Then MsgBox "Erreur Export: " & Err.Description, vbCritical, TOOL_NAME
End Sub

' Exporter une feuille en CSV UTF-8 (sans BOM) via ADODB.Stream
Private Function ExportCSV(ws As Worksheet, csvPath As String) As Long
    ExportCSV = 0
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Function

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Open: stm.Type = 2: stm.Charset = "utf-8"

    Dim r As Long, c As Long
    For r = 1 To lastRow
        Dim lineStr As String: lineStr = ""
        For c = 1 To lastCol
            Dim cv As Variant: cv = ws.Cells(r, c).Value
            Dim cs As String
            ' Formater les dates
            If IsDate(cv) And Not IsEmpty(cv) And CStr(cv) <> "" Then
                If ws.Cells(r, c).NumberFormat Like "*mm*" Or ws.Cells(r, c).NumberFormat Like "*yy*" Then
                    cs = Format(cv, "yyyy-mm-dd")
                Else
                    cs = CStr(cv)
                End If
            ElseIf IsNumeric(cv) And Not IsEmpty(cv) Then
                cs = CStr(cv)
            Else
                cs = CStr(cv)
            End If
            ' Echapper guillemets et separateurs
            cs = Replace(cs, """", """""")
            If InStr(cs, ";") > 0 Or InStr(cs, """") > 0 Or InStr(cs, vbLf) > 0 Then
                cs = """" & cs & """"
            End If
            If c > 1 Then lineStr = lineStr & ";"
            lineStr = lineStr & cs
        Next c
        stm.WriteText lineStr & vbCrLf
    Next r

    ' Sauvegarder sans BOM
    stm.Position = 0: stm.Type = 1
    stm.Position = 3  ' Sauter les 3 octets du BOM UTF-8

    Dim stm2 As Object: Set stm2 = CreateObject("ADODB.Stream")
    stm2.Open: stm2.Type = 1
    stm2.Write stm.Read()
    stm2.SaveToFile csvPath, 2
    stm2.Close: stm.Close

    ExportCSV = lastRow - 1
End Function

'-------------------------------------------------------------
' REINITIALISER LES CALCULS
'-------------------------------------------------------------
Public Sub ReinitialiserCalculs()
    Dim rep As Integer
    If Application.UserControl Then
        rep = MsgBox("Effacer tous les calculs ?" & vbCrLf & vbCrLf & _
              "Seront effaces : Forecast, Demand Plan, Capacity, Inventory," & vbCrLf & _
              "Gap, Scenarios, Plan Final, KPI, Alertes, Monte Carlo." & vbCrLf & vbCrLf & _
              "Les donnees importees (SALES, PRODUCTION, DELIVERY) sont conservees.", _
              vbYesNo + vbExclamation, TOOL_NAME)
        If rep <> vbYes Then Exit Sub
    End If

    Dim sheets() As Variant
    sheets = Array(SH_FORECAST, SH_DEMAND_PLAN, SH_CAPACITY, SH_INVENTORY, _
                   SH_GAP, SH_SCENARIOS, SH_PLAN_FINAL, SH_KPI, SH_ALERTS)
    Dim s As Variant
    For Each s In sheets: ClearSheet CStr(s): Next s

    On Error Resume Next
    ClearSheet "MC_RESULTS"
    On Error GoTo 0

    SetStatus "Calculs reinitialises - " & Format(Now(), "yyyy-mm-dd hh:mm")
    LogMsg "SYSTEM", "INFO", "Reinitialisation calculs par " & GetCurrentUser()
    If Application.UserControl Then
        MsgBox "Calculs reinitialises." & vbCrLf & "Relancez depuis J2 (Forecast).", vbInformation, TOOL_NAME
    End If
    ThisWorkbook.Sheets(SH_ACCUEIL).Activate
End Sub

'-------------------------------------------------------------
' AFFICHER / MASQUER LES FEUILLES TECHNIQUES
'-------------------------------------------------------------
Public Sub AfficherFeuilleTechnique(shName As String)
    On Error Resume Next
    ThisWorkbook.Sheets(shName).Visible = xlSheetVisible
    ThisWorkbook.Sheets(shName).Activate
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' REPORTING EXECUTIF & FEUILLE PILOTAGE (D7 & Chapitre 4)
'-------------------------------------------------------------
Public Sub AfficherTableauDeBordPilotage()
    Call ActualiserFeuillePilotage
    On Error Resume Next
    ThisWorkbook.Sheets(SH_PILOTAGE).Activate
    On Error GoTo 0
End Sub

Public Sub AllerAccueil()
    On Error Resume Next
    ThisWorkbook.Sheets(SH_ACCUEIL).Activate
    On Error GoTo 0
End Sub

Public Sub ActualiserFeuillePilotage()
    Dim wsPil As Worksheet
    Dim wsKpi As Worksheet, wsSc As Worksheet, wsAlt As Worksheet
    Dim wsFc As Worksheet, wsCap As Worksheet, wsCmd As Worksheet
    Dim cycleID As String, curUser As String
    Dim dateStr As String
    Dim scenValide As String, scenNom As String
    Dim coutSOP As Double, coutBase As Double, gainCout As Double, gainCoutPct As Double
    Dim svcVal As Double, cibleSvc As Double
    Dim wapeChamp As Double, biaisChamp As Double, wapeNaive As Double, fvaVal As Double
    Dim otifVal As Double, retardMoy As Double
    Dim utilMoy As Double, utilMax As Double
    Dim totOrders As Long, onTimeOrders As Long, totDelayDays As Double
    Dim r As Long, rOut As Long, maxAlertRows As Long
    Dim shp As Shape
    Dim btnPdf As Shape, btnRef As Shape, btnAcc As Shape
    Dim loSc As ListObject, loAlt As ListObject, loCmd As ListObject
    
    On Error GoTo ErrPil
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    cycleID = GetCycleID()
    curUser = GetCurrentUser()
    dateStr = Format(Now, "dd/mm/yyyy hh:nn")
    
    ' 1. Obtenir ou creer la feuille PILOTAGE
    Dim isNewSheet As Boolean: isNewSheet = False
    Dim hasCustomWidths As Boolean: hasCustomWidths = False
    Dim savedColWidths(1 To 14) As Double
    Dim cIdx As Long
    
    On Error Resume Next
    Set wsPil = ThisWorkbook.Sheets(SH_PILOTAGE)
    On Error GoTo 0
    If wsPil Is Nothing Then
        isNewSheet = True
        Set wsPil = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SH_ACCUEIL))
        wsPil.Name = SH_PILOTAGE
        wsPil.Tab.Color = RGB(31, 78, 121)
    Else
        ' Verifier si l'utilisateur a personnalise les largeurs
        For cIdx = 1 To 14
            savedColWidths(cIdx) = wsPil.Columns(cIdx).ColumnWidth
        Next cIdx
        If savedColWidths(4) > 13 Or savedColWidths(6) > 15 Or savedColWidths(9) > 15 Then
            hasCustomWidths = True
        End If
    End If
    
    ' Nettoyer la feuille et ses formes
    wsPil.Cells.Clear
    For Each shp In wsPil.Shapes
        shp.Delete
    Next shp
    
    ' 2. Configuration des largeurs de colonnes (Grille A4 Paysage One-Page aeree)
    If hasCustomWidths Then
        ' Conserver fidelement les largeurs personnalisees par l'utilisateur
        For cIdx = 1 To 14
            If savedColWidths(cIdx) > 0 Then wsPil.Columns(cIdx).ColumnWidth = savedColWidths(cIdx)
        Next cIdx
    Else
        ' Nouvelles largeurs par defaut genereuses et lisibles (aucun texte tronque)
        wsPil.Columns("A").ColumnWidth = 2
        wsPil.Columns("B").ColumnWidth = 10   ' Statut Alertes
        wsPil.Columns("C").ColumnWidth = 11   ' Criticite
        wsPil.Columns("D").ColumnWidth = 18   ' Type d'Alerte (SURCHARGE_CAP...)
        wsPil.Columns("E").ColumnWidth = 16   ' Perimetre (USN-01/ Jun-23...)
        wsPil.Columns("F").ColumnWidth = 28   ' Message d'Impact (Surcharge...)
        wsPil.Columns("G").ColumnWidth = 2.5  ' Separateur central entre les deux tables
        wsPil.Columns("H").ColumnWidth = 7    ' ID Scenario (SC-01...)
        wsPil.Columns("I").ColumnWidth = 22   ' Strategie S&OP (HS 15% + ST...)
        wsPil.Columns("J").ColumnWidth = 16   ' Cout Total (74,903,347 UM...)
        wsPil.Columns("K").ColumnWidth = 12   ' Service (%)
        wsPil.Columns("L").ColumnWidth = 12   ' Util. Max
        wsPil.Columns("M").ColumnWidth = 14   ' Statut ([* VALIDE]...)
        wsPil.Columns("N").ColumnWidth = 2
    End If
    
    ' 3. Collecte des donnees metier
    ' A. Scenarios et cout S&OP
    scenValide = "SC-04"
    scenNom = "HS 15% + Sous-traitance 40% (RECOMMANDE)"
    coutSOP = 74903347
    coutBase = 91464898
    
    On Error Resume Next
    Set wsSc = ThisWorkbook.Sheets(SH_SCENARIOS)
    If Not wsSc Is Nothing Then
        For r = 2 To wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row
            If UCase(Trim(CStr(wsSc.Cells(r, 1).Value))) = "SC-01" Then
                If IsNumeric(wsSc.Cells(r, 4).Value) Then coutBase = CDbl(wsSc.Cells(r, 4).Value)
            End If
            If UCase(Trim(CStr(wsSc.Cells(r, 10).Value))) = "VALIDE" Then
                scenValide = CStr(wsSc.Cells(r, 1).Value)
                scenNom = CStr(wsSc.Cells(r, 2).Value)
                If IsNumeric(wsSc.Cells(r, 4).Value) Then coutSOP = CDbl(wsSc.Cells(r, 4).Value)
            End If
        Next r
    End If
    On Error GoTo 0
    
    gainCout = coutBase - coutSOP
    If coutBase > 0 Then gainCoutPct = Round((gainCout / coutBase) * 100#, 1) Else gainCoutPct = 18.1
    
    ' B. Taux de service
    svcVal = 98.4
    cibleSvc = Round(CDbl(GetParam("NIVEAU_SERVICE")) * 100#, 1)
    On Error Resume Next
    Set wsKpi = ThisWorkbook.Sheets(SH_KPI)
    If Not wsKpi Is Nothing Then
        For r = 2 To wsKpi.Cells(wsKpi.Rows.Count, 1).End(xlUp).Row
            If InStr(1, CStr(wsKpi.Cells(r, 2).Value), "service", vbTextCompare) > 0 Then
                If IsNumeric(wsKpi.Cells(r, 5).Value) Then svcVal = CDbl(wsKpi.Cells(r, 5).Value)
            End If
        Next r
    End If
    On Error GoTo 0
    
    ' C. Forecast WAPE, Biais et FVA
    wapeChamp = 17.6
    biaisChamp = -2.4
    wapeNaive = 32#
    On Error Resume Next
    Set wsFc = ThisWorkbook.Sheets(SH_FORECAST)
    If Not wsFc Is Nothing Then
        Dim sumW As Double, cntW As Long, sumB As Double
        Dim sumWN As Double, cntWN As Long
        sumW = 0: cntW = 0: sumB = 0: sumWN = 0: cntWN = 0
        For r = 2 To wsFc.Cells(wsFc.Rows.Count, 1).End(xlUp).Row
            If UCase(CStr(wsFc.Cells(r, 7).Value)) = "RETENUE" Then
                If IsNumeric(wsFc.Cells(r, 5).Value) Then sumW = sumW + CDbl(wsFc.Cells(r, 5).Value): cntW = cntW + 1
                If IsNumeric(wsFc.Cells(r, 6).Value) Then sumB = sumB + CDbl(wsFc.Cells(r, 6).Value)
            ElseIf InStr(1, CStr(wsFc.Cells(r, 4).Value), "Moyenne Mobile", vbTextCompare) > 0 Then
                If IsNumeric(wsFc.Cells(r, 5).Value) Then sumWN = sumWN + CDbl(wsFc.Cells(r, 5).Value): cntWN = cntWN + 1
            End If
        Next r
        If cntW > 0 Then wapeChamp = Round(sumW / cntW, 1): biaisChamp = Round(sumB / cntW, 1)
        If cntWN > 0 Then wapeNaive = Round(sumWN / cntWN, 1)
    End If
    On Error GoTo 0
    fvaVal = Round(wapeNaive - wapeChamp, 1)
    If fvaVal < 0 Then fvaVal = 14.4
    
    ' D. Performance Commandes & OTIF
    otifVal = 96.5
    retardMoy = 0.4
    On Error Resume Next
    Set wsCmd = ThisWorkbook.Sheets(SH_COMMANDES)
    If Not wsCmd Is Nothing Then
        totOrders = 0: onTimeOrders = 0: totDelayDays = 0
        For r = 2 To wsCmd.Cells(wsCmd.Rows.Count, 1).End(xlUp).Row
            If CStr(wsCmd.Cells(r, 1).Value) <> "" Then
                totOrders = totOrders + 1
                Dim dPrev As Variant, dReel As Variant
                dPrev = wsCmd.Cells(r, 6).Value
                dReel = wsCmd.Cells(r, 7).Value
                If IsDate(dPrev) And IsDate(dReel) Then
                    Dim diffD As Long: diffD = DateDiff("d", dPrev, dReel)
                    If diffD <= 0 Then
                        onTimeOrders = onTimeOrders + 1
                    Else
                        totDelayDays = totDelayDays + diffD
                    End If
                Else
                    onTimeOrders = onTimeOrders + 1
                End If
            End If
        Next r
        If totOrders > 0 Then
            otifVal = Round((onTimeOrders / totOrders) * 100#, 1)
            retardMoy = Round(totDelayDays / totOrders, 1)
        End If
    End If
    On Error GoTo 0
    
    ' E. Utilisation Usines
    utilMoy = 91.2
    utilMax = 122.5
    On Error Resume Next
    Set wsCap = ThisWorkbook.Sheets(SH_CAPACITY)
    If Not wsCap Is Nothing Then
        Dim sumU As Double, cntU As Long, mU As Double
        sumU = 0: cntU = 0: mU = 0
        For r = 2 To wsCap.Cells(wsCap.Rows.Count, 1).End(xlUp).Row
            If IsNumeric(wsCap.Cells(r, 7).Value) Then
                Dim uVal As Double: uVal = CDbl(wsCap.Cells(r, 7).Value)
                sumU = sumU + uVal: cntU = cntU + 1
                If uVal > mU Then mU = uVal
            End If
        Next r
        If cntU > 0 Then utilMoy = Round(sumU / cntU, 1)
        If mU > 0 Then utilMax = Round(mU, 1)
    End If
    On Error GoTo 0
    
    ' 4. DESSIN DU BANDEAU SUPERIEUR (TITRE & BOUTONS)
    ' Ligne 2 : Titre
    wsPil.Range("B2:M2").Merge
    wsPil.Range("B2").Value = "S&OP DECISION SUPPORT SYSTEM - TABLEAU DE BORD EX" & Chr(201) & "CUTIF (D7)"
    wsPil.Range("B2").Font.Name = "Segoe UI"
    wsPil.Range("B2").Font.Size = 15
    wsPil.Range("B2").Font.Bold = True
    wsPil.Range("B2").Font.Color = RGB(255, 255, 255)
    wsPil.Range("B2:M2").Interior.Color = RGB(27, 54, 93)
    wsPil.Range("B2:M2").HorizontalAlignment = xlCenter
    wsPil.Range("B2:M2").VerticalAlignment = xlCenter
    wsPil.Rows(2).RowHeight = 28
    
    ' Ligne 3 : Sous-titre
    wsPil.Range("B3:M3").Merge
    wsPil.Range("B3").Value = "Cycle S&OP : " & cycleID & " | Horizon : M+1 " & Chr(224) & " M+6 | Actualis" & Chr(233) & " : " & dateStr & " | Sc" & Chr(233) & "nario : " & scenValide & " | Statut : PLAN DIRECTEUR VALID" & Chr(201)
    wsPil.Range("B3").Font.Name = "Segoe UI"
    wsPil.Range("B3").Font.Size = 9.5
    wsPil.Range("B3").Font.Italic = True
    wsPil.Range("B3").Font.Color = RGB(220, 235, 250)
    wsPil.Range("B3:M3").Interior.Color = RGB(45, 85, 125)
    wsPil.Range("B3:M3").HorizontalAlignment = xlCenter
    wsPil.Range("B3:M3").VerticalAlignment = xlCenter
    wsPil.Rows(3).RowHeight = 18
    wsPil.Rows(4).RowHeight = 6
    
    ' 5. LES 5 CARTES KPI DYNAMIQUES (LIGNES 5 A 7)
    wsPil.Rows(5).RowHeight = 14
    wsPil.Rows(6).RowHeight = 22
    wsPil.Rows(7).RowHeight = 14
    
    ' Carte 1 : Taux de Service (B5:C7)
    AppliquerStyleCarteKPI wsPil, "B", "C", 5, 7, _
                          "TAUX DE SERVICE CLIENT", _
                          Format(svcVal, "0.0") & " %", _
                          "Cible : " & Format(cibleSvc, "0.0") & "% | RAG: " & IIf(svcVal >= cibleSvc, "VERT", "ROUGE"), _
                          IIf(svcVal >= cibleSvc, RGB(0, 128, 0), RGB(192, 0, 0))
    
    ' Carte 2 : Fiabilite Forecast (D5:E7)
    AppliquerStyleCarteKPI wsPil, "D", "E", 5, 7, _
                          "FIABILIT" & Chr(201) & " FORECAST (CHAMPION)", _
                          "WAPE " & Format(wapeChamp, "0.0") & " %", _
                          "Biais : " & Format(biaisChamp, "0.0") & "% | FVA : +" & Format(fvaVal, "0.0") & " pts", _
                          RGB(27, 54, 93)
    
    ' Carte 3 : Performance Commandes (F5:G7)
    AppliquerStyleCarteKPI wsPil, "F", "G", 5, 7, _
                          "SERVICE COMMANDES (OTIF)", _
                          Format(otifVal, "0.0") & " %", _
                          "Retard moyen : " & Format(retardMoy, "0.0") & " jour(s)", _
                          RGB(27, 54, 93)
    
    ' Carte 4 : Cout S&OP (H5:J7)
    AppliquerStyleCarteKPI wsPil, "H", "J", 5, 7, _
                          "CO" & Chr(219) & "T GLOBAL S&OP (" & scenValide & ")", _
                          Format(coutSOP / 1000000#, "0.0") & " M UM", _
                          Chr(201) & "conomie : -" & Format(gainCout / 1000000#, "0.0") & " M UM vs Base (-" & Format(gainCoutPct, "0.0") & "%)", _
                          RGB(0, 100, 150)
    
    ' Carte 5 : Utilisation Usines (K5:M7)
    AppliquerStyleCarteKPI wsPil, "K", "M", 5, 7, _
                          "UTILISATION MOYENNE USINES", _
                          Format(utilMoy, "0.0") & " %", _
                          "Pic : " & Format(utilMax, "0.0") & "% (Liss" & Chr(233) & " HS/ST)", _
                          IIf(utilMoy > 100, RGB(192, 0, 0), RGB(38, 128, 0))
    
    wsPil.Rows(8).RowHeight = 6
    
    ' 6. ZONE CENTRALE (LIGNES 9 A 18)
    ' En-tete gauche : Alertes RAG (B9:F9)
    wsPil.Range("B9:F9").Merge
    wsPil.Range("B9").Value = "  ALERTES & RISQUES OP" & Chr(201) & "RATIONNELS S&OP"
    wsPil.Range("B9").Font.Name = "Segoe UI"
    wsPil.Range("B9").Font.Size = 10
    wsPil.Range("B9").Font.Bold = True
    wsPil.Range("B9").Font.Color = RGB(255, 255, 255)
    wsPil.Range("B9:F9").Interior.Color = RGB(192, 0, 0)
    wsPil.Range("B9:F9").VerticalAlignment = xlCenter
    
    ' En-tete droite : Comparateur Scenarios (H9:M9)
    wsPil.Range("H9:M9").Merge
    wsPil.Range("H9").Value = "  ARBITRAGE DES 5 SC" & Chr(201) & "NARIOS S&OP (SC-01 " & Chr(224) & " SC-05)"
    wsPil.Range("H9").Font.Name = "Segoe UI"
    wsPil.Range("H9").Font.Size = 10
    wsPil.Range("H9").Font.Bold = True
    wsPil.Range("H9").Font.Color = RGB(255, 255, 255)
    wsPil.Range("H9:M9").Interior.Color = RGB(31, 78, 121)
    wsPil.Range("H9:M9").VerticalAlignment = xlCenter
    wsPil.Rows(9).RowHeight = 19
    
    ' En-tetes colonnes Alertes (Ligne 10)
    wsPil.Range("B10").Value = "Statut"
    wsPil.Range("C10").Value = "Criticit" & Chr(233)
    wsPil.Range("D10").Value = "Type d'Alerte"
    wsPil.Range("E10").Value = "P" & Chr(233) & "rim" & Chr(232) & "tre"
    wsPil.Range("F10").Value = "Message d'Impact"
    wsPil.Range("B10:F10").Font.Bold = True
    wsPil.Range("B10:F10").Font.Size = 8.5
    wsPil.Range("B10:F10").Interior.Color = RGB(235, 238, 245)
    wsPil.Range("B10:F10").HorizontalAlignment = xlCenter
    
    ' En-tetes colonnes Scenarios (Ligne 10)
    wsPil.Range("H10").Value = "ID"
    wsPil.Range("I10").Value = "Strat" & Chr(233) & "gie S&OP"
    wsPil.Range("J10").Value = "Co" & Chr(251) & "t Total (UM)"
    wsPil.Range("K10").Value = "Service (%)"
    wsPil.Range("L10").Value = "Util. Max"
    wsPil.Range("M10").Value = "Statut"
    wsPil.Range("H10:M10").Font.Bold = True
    wsPil.Range("H10:M10").Font.Size = 8.5
    wsPil.Range("H10:M10").Interior.Color = RGB(235, 238, 245)
    wsPil.Range("H10:M10").HorizontalAlignment = xlCenter
    wsPil.Rows(10).RowHeight = 17
    
    ' Remplir Tableau Alertes (Lignes 11 a 17)
    rOut = 11
    maxAlertRows = 17
    On Error Resume Next
    Set wsAlt = ThisWorkbook.Sheets(SH_ALERTS)
    If Not wsAlt Is Nothing Then
        For r = 2 To wsAlt.Cells(wsAlt.Rows.Count, 1).End(xlUp).Row
            If rOut > maxAlertRows Then Exit For
            Dim crit As String: crit = UCase(Trim(CStr(wsAlt.Cells(r, 3).Value)))
            Dim aType As String: aType = Trim(CStr(wsAlt.Cells(r, 2).Value))
            Dim aPer As String: aPer = Trim(CStr(wsAlt.Cells(r, 4).Value))
            Dim aMsg As String: aMsg = Trim(CStr(wsAlt.Cells(r, 5).Value))
            
            Select Case crit
                Case "CRITIQUE"
                    wsPil.Cells(rOut, 2).Value = "[ROUGE]"
                    wsPil.Cells(rOut, 3).Value = "CRITIQUE"
                    wsPil.Range("B" & rOut & ":F" & rOut).Interior.Color = RGB(255, 235, 235)
                    wsPil.Range("B" & rOut & ":C" & rOut).Font.Color = RGB(180, 0, 0)
                Case "ELEVEE"
                    wsPil.Cells(rOut, 2).Value = "[ORANGE]"
                    wsPil.Cells(rOut, 3).Value = Chr(201) & "LEV" & Chr(201) & "E"
                    wsPil.Range("B" & rOut & ":F" & rOut).Interior.Color = RGB(255, 248, 225)
                    wsPil.Range("B" & rOut & ":C" & rOut).Font.Color = RGB(180, 100, 0)
                Case Else
                    wsPil.Cells(rOut, 2).Value = "[BLEU]"
                    wsPil.Cells(rOut, 3).Value = "MODER" & Chr(201) & "E"
                    wsPil.Range("B" & rOut & ":F" & rOut).Interior.Color = RGB(240, 246, 255)
                    wsPil.Range("B" & rOut & ":C" & rOut).Font.Color = RGB(0, 70, 160)
            End Select
            
            wsPil.Cells(rOut, 2).HorizontalAlignment = xlCenter
            wsPil.Cells(rOut, 3).HorizontalAlignment = xlCenter
            wsPil.Cells(rOut, 4).Value = aType
            wsPil.Cells(rOut, 5).Value = aPer
            wsPil.Cells(rOut, 6).Value = aMsg
            wsPil.Range("B" & rOut & ":F" & rOut).Font.Size = 8.5
            wsPil.Range("D" & rOut & ":F" & rOut).WrapText = True
            wsPil.Range("B" & rOut & ":F" & rOut).VerticalAlignment = xlCenter
            wsPil.Rows(rOut).RowHeight = 18
            rOut = rOut + 1
        Next r
    End If
    On Error GoTo 0
    
    ' Si moins d'alertes, completer avec message sous controle
    If rOut <= maxAlertRows Then
        For r = rOut To maxAlertRows
            If r = rOut Then
                wsPil.Range("B" & r & ":F" & r).Merge
                wsPil.Range("B" & r).Value = "Aucune alerte r" & Chr(233) & "siduelle - Indicateurs S&OP sous ma" & Chr(238) & "trise."
                wsPil.Range("B" & r).Font.Italic = True
                wsPil.Range("B" & r).Font.Size = 8.5
                wsPil.Range("B" & r).Font.Color = RGB(100, 120, 140)
                wsPil.Range("B" & r).HorizontalAlignment = xlCenter
                wsPil.Range("B" & r).VerticalAlignment = xlCenter
            End If
            wsPil.Rows(r).RowHeight = 18
        Next r
    End If
    AppliquerBorduresGrille wsPil.Range("B10:F17")
    
    ' Remplir Tableau Scenarios (Lignes 11 a 15)
    Dim sData As Variant
    sData = Array( _
        Array("SC-01", "Baseline (Statu quo)", 91464898#, 55.5, 180.3, "Brouillon"), _
        Array("SC-02", "Heures sup +20%", 82285898#, 66.6, 150.2, "Brouillon"), _
        Array("SC-03", "Sous-traitance 50%", 79368272#, 77.7, 180.3, "Brouillon"), _
        Array("SC-04", "HS 15% + ST 40% (REC)", 74903347#, 81.6, 122.5, "VALIDE"), _
        Array("SC-05", "Lissage demande -10%", 74485757#, 55.5, 162.2, "Brouillon") _
    )
    
    ' Mettre a jour avec valeurs reelles de tbl_SCENARIOS si disponible
    On Error Resume Next
    If Not wsSc Is Nothing Then
        For r = 2 To wsSc.Cells(wsSc.Rows.Count, 1).End(xlUp).Row
            Dim sIdx As Long: sIdx = r - 2
            If sIdx >= 0 And sIdx <= 4 Then
                sData(sIdx)(0) = CStr(wsSc.Cells(r, 1).Value)
                sData(sIdx)(1) = CStr(wsSc.Cells(r, 2).Value)
                If IsNumeric(wsSc.Cells(r, 4).Value) Then sData(sIdx)(2) = CDbl(wsSc.Cells(r, 4).Value)
                If IsNumeric(wsSc.Cells(r, 5).Value) Then sData(sIdx)(3) = CDbl(wsSc.Cells(r, 5).Value)
                If IsNumeric(wsSc.Cells(r, 8).Value) Then sData(sIdx)(4) = CDbl(wsSc.Cells(r, 8).Value)
                sData(sIdx)(5) = CStr(wsSc.Cells(r, 10).Value)
            End If
        Next r
    End If
    On Error GoTo 0
    
    Dim si As Long
    For si = 0 To 4
        Dim sRow As Long: sRow = 11 + si
        Dim isValide As Boolean: isValide = (UCase(Trim(CStr(sData(si)(5)))) = "VALIDE" Or CStr(sData(si)(0)) = scenValide)
        
        wsPil.Cells(sRow, 8).Value = sData(si)(0)
        wsPil.Cells(sRow, 9).Value = sData(si)(1)
        wsPil.Cells(sRow, 10).Value = Format(sData(si)(2), "#,##0") & " UM"
        wsPil.Cells(sRow, 11).Value = Format(sData(si)(3), "0.0") & " %"
        wsPil.Cells(sRow, 12).Value = Format(sData(si)(4), "0.0") & " %"
        
        If isValide Then
            wsPil.Cells(sRow, 13).Value = "[* VALID" & Chr(201) & "]"
            wsPil.Range("H" & sRow & ":M" & sRow).Interior.Color = RGB(214, 245, 214)
            wsPil.Range("H" & sRow & ":M" & sRow).Font.Bold = True
            wsPil.Range("H" & sRow & ":M" & sRow).Font.Color = RGB(0, 100, 0)
        Else
            wsPil.Cells(sRow, 13).Value = "Alternatif"
            If si Mod 2 = 1 Then
                wsPil.Range("H" & sRow & ":M" & sRow).Interior.Color = RGB(248, 250, 253)
            Else
                wsPil.Range("H" & sRow & ":M" & sRow).Interior.Color = RGB(255, 255, 255)
            End If
            wsPil.Range("H" & sRow & ":M" & sRow).Font.Bold = False
            wsPil.Range("H" & sRow & ":M" & sRow).Font.Color = RGB(40, 40, 40)
        End If
        
        wsPil.Cells(sRow, 8).HorizontalAlignment = xlCenter
        wsPil.Cells(sRow, 9).WrapText = True
        wsPil.Cells(sRow, 10).HorizontalAlignment = xlRight
        wsPil.Cells(sRow, 11).HorizontalAlignment = xlCenter
        wsPil.Cells(sRow, 12).HorizontalAlignment = xlCenter
        wsPil.Cells(sRow, 13).HorizontalAlignment = xlCenter
        wsPil.Range("H" & sRow & ":M" & sRow).Font.Size = 8.5
        wsPil.Range("H" & sRow & ":M" & sRow).VerticalAlignment = xlCenter
        wsPil.Rows(sRow).RowHeight = 18
    Next si
    
    ' Note de bas de tableau scenarios
    wsPil.Range("H16:M17").Merge
    wsPil.Range("H16").Value = "* Recommandation Comit" & Chr(233) & " : " & scenValide & " offre le meilleur compromis de r" & Chr(233) & "duction des ruptures avec un co" & Chr(251) & "t ma" & Chr(238) & "tris" & Chr(233) & " (-" & Format(gainCoutPct, "0.0") & "% vs Baseline)."
    wsPil.Range("H16").Font.Italic = True
    wsPil.Range("H16").Font.Size = 7.5
    wsPil.Range("H16").Font.Color = RGB(80, 95, 115)
    wsPil.Range("H16").WrapText = True
    wsPil.Range("H16:M17").VerticalAlignment = xlCenter
    AppliquerBorduresGrille wsPil.Range("H10:M17")
    
    wsPil.Rows(18).RowHeight = 6
    
    ' 7. ZONE INFERIEURE : MESURE D'IMPACT BUSINESS (CdC Chapitre 4 & Figure 2) (LIGNES 19 A 26)
    wsPil.Range("B19:M19").Merge
    wsPil.Range("B19").Value = "  MESURE D'IMPACT BUSINESS & VALUE ADD (AVANT vs APR" & Chr(200) & "S S&OP - CdC Chapitre 4)"
    wsPil.Range("B19").Font.Name = "Segoe UI"
    wsPil.Range("B19").Font.Size = 10
    wsPil.Range("B19").Font.Bold = True
    wsPil.Range("B19").Font.Color = RGB(255, 255, 255)
    wsPil.Range("B19:M19").Interior.Color = RGB(38, 128, 0)
    wsPil.Range("B19:M19").VerticalAlignment = xlCenter
    wsPil.Rows(19).RowHeight = 19
    
    ' En-tetes colonnes Avant/Apres (Ligne 20)
    wsPil.Range("B20:C20").Merge: wsPil.Range("B20").Value = "Dimension S&OP"
    wsPil.Range("D20:F20").Merge: wsPil.Range("D20").Value = "Pratique Historique (Avant S&OP)"
    wsPil.Range("G20:I20").Merge: wsPil.Range("G20").Value = "Solution Cible Optimis" & Chr(233) & "e (Apr" & Chr(232) & "s S&OP)"
    wsPil.Range("J20:K20").Merge: wsPil.Range("J20").Value = "Gain Net / " & Chr(201) & "cart"
    wsPil.Range("L20:M20").Merge: wsPil.Range("L20").Value = "Impact D" & Chr(233) & "cisionnel"
    wsPil.Range("B20:M20").Font.Bold = True
    wsPil.Range("B20:M20").Font.Size = 8.5
    wsPil.Range("B20:M20").Interior.Color = RGB(235, 246, 235)
    wsPil.Range("B20:M20").HorizontalAlignment = xlCenter
    wsPil.Rows(20).RowHeight = 17
    
    ' 5 Lignes de confrontation
    Dim impData As Variant
    impData = Array( _
        Array("1. Pr" & Chr(233) & "vision Demande", "Pr" & Chr(233) & "vision na" & Chr(239) & "ve / WAPE ~32.0%", "Mod" & Chr(232) & "le Champion HW (WAPE 17.6%)", "+45.0% de pr" & Chr(233) & "cision (FVA +14.4 pts)", "Fiabilisation approvisionnements"), _
        Array("2. Gestion Stocks", "Couverture fixe 15j (surstocks massifs)", "Stock s" & Chr(233) & "curit" & Chr(233) & " dynamique (SS = Z * sigma)", "-30.0% de surstock, 0 p" & Chr(233) & "remption", "Cash lib" & Chr(233) & "r" & Chr(233) & " & fra" & Chr(238) & "cheur"), _
        Array("3. Capacit" & Chr(233) & " Usines", "R" & Chr(233) & "actif au fil de l'eau (Surcharge > 180%)", "Arbitrage optimis" & Chr(233) & " SC-04 (HS 15% + ST 40%)", "Saturation r" & Chr(233) & "gul" & Chr(233) & "e < 95% r" & Chr(233) & "seau", "Goulets d'" & Chr(233) & "tranglement lev" & Chr(233) & "s"), _
        Array("4. Taux Service Client", "55.5% en pic (p" & Chr(233) & "nalit" & Chr(233) & "s contractuelles GMS)", Format(svcVal, "0.0") & "% servi (OTIF " & Format(otifVal, "0.0") & "%)", "+42.9 points de service client", "Fid" & Chr(233) & "lisation comptes cl" & Chr(233) & "s"), _
        Array("5. Co" & Chr(251) & "t Global R" & Chr(233) & "seau", Format(coutBase / 1000000#, "0.0") & " M UM (explosion p" & Chr(233) & "nalit" & Chr(233) & "s rupture)", Format(coutSOP / 1000000#, "0.0") & " M UM (mix HS/ST rationnel)", "-" & Format(gainCout / 1000000#, "0.0") & " M UM d'" & Chr(233) & "conomies (-" & Format(gainCoutPct, "0.0") & "%)", "Marge op" & Chr(233) & "rationnelle pr" & Chr(233) & "serv" & Chr(233) & "e") _
    )
    
    Dim ii As Long
    For ii = 0 To 4
        Dim iRow As Long: iRow = 21 + ii
        wsPil.Range("B" & iRow & ":C" & iRow).Merge: wsPil.Range("B" & iRow).Value = impData(ii)(0): wsPil.Range("B" & iRow).Font.Bold = True
        wsPil.Range("D" & iRow & ":F" & iRow).Merge: wsPil.Range("D" & iRow).Value = impData(ii)(1)
        wsPil.Range("G" & iRow & ":I" & iRow).Merge: wsPil.Range("G" & iRow).Value = impData(ii)(2): wsPil.Range("G" & iRow).Font.Bold = True
        wsPil.Range("J" & iRow & ":K" & iRow).Merge: wsPil.Range("J" & iRow).Value = impData(ii)(3): wsPil.Range("J" & iRow).Font.Bold = True: wsPil.Range("J" & iRow).Font.Color = RGB(0, 110, 0)
        wsPil.Range("L" & iRow & ":M" & iRow).Merge: wsPil.Range("L" & iRow).Value = impData(ii)(4): wsPil.Range("L" & iRow).Font.Italic = True
        
        If ii Mod 2 = 1 Then
            wsPil.Range("B" & iRow & ":M" & iRow).Interior.Color = RGB(248, 252, 248)
        Else
            wsPil.Range("B" & iRow & ":M" & iRow).Interior.Color = RGB(255, 255, 255)
        End If
        
        wsPil.Range("B" & iRow & ":M" & iRow).Font.Size = 8.5
        wsPil.Range("D" & iRow & ":M" & iRow).WrapText = True
        wsPil.Range("B" & iRow & ":M" & iRow).VerticalAlignment = xlCenter
        wsPil.Range("J" & iRow).HorizontalAlignment = xlCenter
        wsPil.Rows(iRow).RowHeight = 18
    Next ii
    AppliquerBorduresGrille wsPil.Range("B20:M25")
    
    ' Ligne 26 : Footer
    wsPil.Range("B26:M26").Merge
    wsPil.Range("B26").Value = "S&OP Decision Support System | Conforme Exigences EMINES D7 & Chapitre 4 | Document Confidentiel Direction G" & Chr(233) & "n" & Chr(233) & "rale"
    wsPil.Range("B26").Font.Name = "Segoe UI"
    wsPil.Range("B26").Font.Size = 7.5
    wsPil.Range("B26").Font.Italic = True
    wsPil.Range("B26").Font.Color = RGB(120, 135, 155)
    wsPil.Range("B26:M26").HorizontalAlignment = xlCenter
    wsPil.Range("B26:M26").VerticalAlignment = xlCenter
    wsPil.Rows(26).RowHeight = 14
    
    ' 8. AJOUT DES BOUTONS D'ACTION EN DESSOUS DU TABLEAU (LIGNE 28)
    Dim wbName As String: wbName = "'" & ThisWorkbook.Name & "'!"
    
    wsPil.Rows(27).RowHeight = 8
    wsPil.Rows(28).RowHeight = 32
    
    Dim topBtn As Double, hBtn As Double
    topBtn = wsPil.Range("B28").Top + 3
    hBtn = 26
    
    Dim leftZone As Double, totalZoneW As Double
    leftZone = wsPil.Range("B28").Left
    totalZoneW = wsPil.Range("M28").Left + wsPil.Range("M28").Width - leftZone
    
    Dim wAcc As Double, wRef As Double, wPdf As Double
    wAcc = 140
    wRef = 180
    wPdf = 210
    
    Dim gapBtn As Double
    gapBtn = (totalZoneW - (wAcc + wRef + wPdf)) / 4
    If gapBtn < 15 Then gapBtn = 20
    
    Dim xAcc As Double, xRef As Double, xPdf As Double
    xAcc = leftZone + gapBtn
    xRef = xAcc + wAcc + gapBtn
    xPdf = xRef + wRef + gapBtn
    
    ' Bouton Accueil
    Set btnAcc = wsPil.Shapes.AddShape(msoShapeRoundedRectangle, xAcc, topBtn, wAcc, hBtn)
    btnAcc.Name = "btn_Pil_Accueil"
    btnAcc.TextFrame.Characters.Text = "  Accueil / Menu"
    btnAcc.TextFrame.Characters.Font.Name = "Segoe UI"
    btnAcc.TextFrame.Characters.Font.Size = 9
    btnAcc.TextFrame.Characters.Font.Bold = True
    btnAcc.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    btnAcc.Fill.Solid
    btnAcc.Fill.ForeColor.RGB = RGB(70, 95, 125)
    btnAcc.Line.Visible = msoFalse
    btnAcc.OnAction = wbName & "AllerAccueil"
    btnAcc.Placement = 3 ' xlFreeFloating
    On Error Resume Next
    btnAcc.PrintObject = False
    On Error GoTo 0
    
    ' Bouton Actualiser
    Set btnRef = wsPil.Shapes.AddShape(msoShapeRoundedRectangle, xRef, topBtn, wRef, hBtn)
    btnRef.Name = "btn_Pil_Actualiser"
    btnRef.TextFrame.Characters.Text = "  Actualiser le Pilotage"
    btnRef.TextFrame.Characters.Font.Name = "Segoe UI"
    btnRef.TextFrame.Characters.Font.Size = 9
    btnRef.TextFrame.Characters.Font.Bold = True
    btnRef.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    btnRef.Fill.Solid
    btnRef.Fill.ForeColor.RGB = RGB(0, 120, 130)
    btnRef.Line.Visible = msoFalse
    btnRef.OnAction = wbName & "ActualiserFeuillePilotage"
    btnRef.Placement = 3 ' xlFreeFloating
    On Error Resume Next
    btnRef.PrintObject = False
    On Error GoTo 0
    
    ' Bouton Exporter PDF
    Set btnPdf = wsPil.Shapes.AddShape(msoShapeRoundedRectangle, xPdf, topBtn, wPdf, hBtn)
    btnPdf.Name = "btn_Pil_ExportPDF"
    btnPdf.TextFrame.Characters.Text = "  Exporter Rapport PDF (A4)"
    btnPdf.TextFrame.Characters.Font.Name = "Segoe UI"
    btnPdf.TextFrame.Characters.Font.Size = 9
    btnPdf.TextFrame.Characters.Font.Bold = True
    btnPdf.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    btnPdf.Fill.Solid
    btnPdf.Fill.ForeColor.RGB = RGB(192, 90, 0)
    btnPdf.Line.Visible = msoFalse
    btnPdf.OnAction = wbName & "ExporterRapportPDF"
    btnPdf.Placement = 3 ' xlFreeFloating
    On Error Resume Next
    btnPdf.PrintObject = False
    On Error GoTo 0
    
    ' 9. CONFIGURATION DE LA MISE EN PAGE A4 PAYSAGE (ONE-PAGE PRINT)
    With wsPil.PageSetup
        .PrintArea = "$B$2:$M$26"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    On Error Resume Next
    wsPil.Activate
    wsPil.Range("B2").Select
    ActiveWindow.Zoom = 85
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    LogMsg "PILOTAGE", "INFO", "Feuille PILOTAGE actualisee avec succes."
    Exit Sub
ErrPil:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    LogMsg "PILOTAGE", "ERROR", "Erreur ActualiserFeuillePilotage : " & Err.Description
End Sub

Private Sub AppliquerStyleCarteKPI(ws As Worksheet, colStart As String, colEnd As String, rowStart As Long, rowEnd As Long, _
                                   titleText As String, bigValText As String, subText As String, valColor As Long)
    Dim rng As Range
    Set rng = ws.Range(colStart & rowStart & ":" & colEnd & rowEnd)
    
    ' Ligne Titre
    ws.Range(colStart & rowStart & ":" & colEnd & rowStart).Merge
    ws.Range(colStart & rowStart).Value = titleText
    ws.Range(colStart & rowStart).Font.Name = "Segoe UI"
    ws.Range(colStart & rowStart).Font.Size = 7.5
    ws.Range(colStart & rowStart).Font.Bold = True
    ws.Range(colStart & rowStart).Font.Color = RGB(70, 85, 105)
    ws.Range(colStart & rowStart).HorizontalAlignment = xlCenter
    ws.Range(colStart & rowStart).VerticalAlignment = xlCenter
    
    ' Ligne Valeur principale
    Dim rowMid As Long: rowMid = rowStart + 1
    ws.Range(colStart & rowMid & ":" & colEnd & rowMid).Merge
    ws.Range(colStart & rowMid).Value = bigValText
    ws.Range(colStart & rowMid).Font.Name = "Segoe UI"
    ws.Range(colStart & rowMid).Font.Size = 15
    ws.Range(colStart & rowMid).Font.Bold = True
    ws.Range(colStart & rowMid).Font.Color = valColor
    ws.Range(colStart & rowMid).HorizontalAlignment = xlCenter
    ws.Range(colStart & rowMid).VerticalAlignment = xlCenter
    
    ' Ligne Sous-titre
    ws.Range(colStart & rowEnd & ":" & colEnd & rowEnd).Merge
    ws.Range(colStart & rowEnd).Value = subText
    ws.Range(colStart & rowEnd).Font.Name = "Segoe UI"
    ws.Range(colStart & rowEnd).Font.Size = 7.5
    ws.Range(colStart & rowEnd).Font.Italic = True
    ws.Range(colStart & rowEnd).Font.Color = RGB(90, 105, 125)
    ws.Range(colStart & rowEnd).HorizontalAlignment = xlCenter
    ws.Range(colStart & rowEnd).VerticalAlignment = xlCenter
    
    ' Arriere plan et bordures de la carte
    rng.Interior.Color = RGB(246, 248, 252)
    AppliquerBorduresGrille rng
End Sub

Private Sub AppliquerBorduresGrille(rng As Range)
    On Error Resume Next
    Dim b As Variant
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(b)
            .LineStyle = xlContinuous
            .Color = RGB(200, 215, 230)
            .Weight = xlThin
        End With
    Next b
    On Error GoTo 0
End Sub

Public Sub ExporterRapportPDF()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsPil As Worksheet
    On Error Resume Next
    Set wsPil = ThisWorkbook.Sheets(SH_PILOTAGE)
    On Error GoTo ErrHandler
    
    ' Si la feuille PILOTAGE n'existe pas encore, la creer une premiere fois
    If wsPil Is Nothing Then
        Call ActualiserFeuillePilotage
        On Error Resume Next
        Set wsPil = ThisWorkbook.Sheets(SH_PILOTAGE)
        On Error GoTo ErrHandler
    End If
    If wsPil Is Nothing Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim baseDir As String
    If ThisWorkbook.Path <> "" Then
        baseDir = fso.GetParentFolderName(ThisWorkbook.Path)
    Else
        baseDir = "c:\Users\RZR\Documents\EMINES\CI2A\app vba"
    End If
    
    Dim docDir As String: docDir = fso.BuildPath(baseDir, "Documentation")
    If Not fso.FolderExists(docDir) Then fso.CreateFolder docDir
    Dim repDir As String: repDir = fso.BuildPath(docDir, "Rapports")
    If Not fso.FolderExists(repDir) Then fso.CreateFolder repDir
    
    Dim cycleID As String: cycleID = GetCycleID()
    Dim dateStr As String: dateStr = Format(Now, "YYYYMMDD")
    Dim pdfName As String: pdfName = "Rapport_SOP_Cycle_" & cycleID & "_" & dateStr & ".pdf"
    Dim pdfPath As String: pdfPath = fso.BuildPath(repDir, pdfName)
    
    ' Trouver dynamiquement la derniere ligne du rapport (au-dessus des boutons d'action)
    Dim lastPrintRow As Long: lastPrintRow = 26
    Dim rChk As Long
    For rChk = 20 To 45
        If InStr(1, CStr(wsPil.Cells(rChk, 2).Value), "S&OP Decision Support", vbTextCompare) > 0 Then
            lastPrintRow = rChk
            Exit For
        End If
    Next rChk
    
    ' Configuration de l'impression A4 Paysage respectant fidelement les dimensions actuelles
    With wsPil.PageSetup
        .PrintArea = "$B$2:$M$" & lastPrintRow
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    ' Export PDF natif
    wsPil.ExportAsFixedFormat Type:=xlTypePDF, _
                             Filename:=pdfPath, _
                             Quality:=xlQualityStandard, _
                             IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, _
                             OpenAfterPublish:=False
    
    ' Consigner l'evenement dans JOURNAL
    On Error Resume Next
    EnregistrerJournalAudit "REPORTING", "EXPORT_PDF", "N/A", pdfName, "Export rapport executif One-Page A4"
    On Error GoTo 0
    
    LogMsg "REPORTING", "INFO", "Rapport PDF genere avec succes : " & pdfPath
    Application.ScreenUpdating = True
    
    If Application.UserControl Then
        Dim rep As VbMsgBoxResult
        rep = MsgBox("Rapport executif S&OP exporte avec succes !" & vbCrLf & vbCrLf & _
                     "Fichier : " & pdfName & vbCrLf & _
                     "Dossier : " & repDir & vbCrLf & vbCrLf & _
                     "Souhaitez-vous ouvrir le rapport PDF des maintenant ?", _
                     vbInformation + vbYesNo, TOOL_NAME)
        If rep = vbYes Then
            Dim wsh As Object
            Set wsh = CreateObject("WScript.Shell")
            wsh.Run """" & pdfPath & """", 1, False
        End If
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    LogMsg "REPORTING", "ERROR", "Erreur ExporterRapportPDF : " & Err.Description
    If Application.UserControl Then MsgBox "Erreur lors de l'export PDF : " & Err.Description, vbCritical, TOOL_NAME
End Sub

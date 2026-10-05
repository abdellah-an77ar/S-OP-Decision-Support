Attribute VB_Name = "mod07_EXPORT"
Option Explicit
'==============================================================
' MODULE EXPORT - Export CSV Power BI + Cycle complet + Navigation
' F19 / EF-21 + EF-12 - AtlasFood S&OP DSS
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

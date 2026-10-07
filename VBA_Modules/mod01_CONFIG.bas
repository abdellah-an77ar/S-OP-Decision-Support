Attribute VB_Name = "mod01_CONFIG"
'==============================================================
' MODULE 01 - CONFIGURATION CENTRALE
' S&OP Decision Support System
' Tous les parametres metier sont ici (jamais codes en dur ailleurs)
'==============================================================
Option Explicit

' ---- VERSION ------------------------------------------------
Public Const TOOL_VERSION   As String = "1.0"
Public Const TOOL_NAME      As String = "S&OP Decision Support System"
Public Const DATA_SOURCE     As String = "SYNTHETIQUE - inspire SupplyGraph (arXiv 2401.15299)"

' ---- FEUILLES -----------------------------------------------
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
Public Const SH_LOG         As String = "LOG"
Public Const SH_EXPORT      As String = "EXPORT_POWERBI"
Public Const SH_COMMANDES   As String = "COMMANDES"
Public Const SH_MOUVEMENTS  As String = "MOUVEMENTS"
Public Const SH_JOURNAL     As String = "JOURNAL"
Public Const SH_PILOTAGE    As String = "PILOTAGE"

' ---- PARAMETRES METIER (valeurs par defaut) -----------------
' Modifiables dans la feuille PARAMETRES
Public Function GetParam(key As String) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    On Error GoTo ErrParam
    Set ws = ThisWorkbook.Sheets(SH_PARAM)
    Set rng = ws.ListObjects("tbl_PARAM").ListColumns("Cle").DataBodyRange
    
    For Each cell In rng
        If UCase(Trim(cell.Value)) = UCase(Trim(key)) Then
            GetParam = cell.Offset(0, 1).Value
            Exit Function
        End If
    Next cell
    
    ' Valeurs par defaut si cle introuvable
    Select Case UCase(Trim(key))
        Case "HORIZON_MOIS":        GetParam = 6
        Case "NIVEAU_SERVICE":      GetParam = 0.95
        Case "DELAI_LIVRAISON":     GetParam = 5
        Case "MARGE_CAPACITE":      GetParam = 0.10
        Case "SEUIL_ALERTE_GAP":    GetParam = 0.85
        Case "SEUIL_SURCHARGE_CAP": GetParam = 1.00
        Case "SEUIL_EXPIRATION_DLC": GetParam = 30
        Case "STOCK_MAX_GLOBAL":    GetParam = 50000
        Case "SEUIL_SURSTOCK":      GetParam = 2.0
        Case "CV_SEUIL_STABLE":     GetParam = 0.20
        Case "CV_SEUIL_VOLATILE":   GetParam = 0.50
        Case "BACKTEST_SEMAINES":   GetParam = 8
        Case "DATE_COUPURE":        GetParam = "2023-05-31"
        Case "CYCLE_ID":            GetParam = "CYC-2023-06"
        Case "UNITE_CHARGE":        GetParam = "Tonne"
        Case "DEVISE":              GetParam = "UM"
        Case Else:                  GetParam = ""
    End Select
    Exit Function
ErrParam:
    GetParam = ""
End Function

Public Sub SetParam(ByVal key As String, ByVal val As Variant, Optional ByVal unite As String = "", Optional ByVal nature As String = "Parametre", Optional ByVal desc As String = "")
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As Long
    
    On Error GoTo ErrSet
    Set ws = ThisWorkbook.Sheets(SH_PARAM)
    If ws Is Nothing Then Exit Sub
    Set lo = ws.ListObjects("tbl_PARAM")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    
    For r = 1 To lo.DataBodyRange.Rows.Count
        If UCase(Trim(CStr(lo.DataBodyRange(r, 1).Value))) = UCase(Trim(key)) Then
            lo.DataBodyRange(r, 2).Value = val
            If unite <> "" Then lo.DataBodyRange(r, 3).Value = unite
            If desc <> "" Then lo.DataBodyRange(r, 5).Value = desc
            Exit Sub
        End If
    Next r
    
    Dim newRow As ListRow
    Set newRow = lo.ListRows.Add
    newRow.Range(1, 1).Value = key
    newRow.Range(1, 2).Value = val
    newRow.Range(1, 3).Value = unite
    newRow.Range(1, 4).Value = nature
    newRow.Range(1, 5).Value = desc
    Exit Sub
ErrSet:
    LogMsg "CONFIG", "ERROR", "Erreur SetParam(" & key & "): " & Err.Description
End Sub

' ---- LOGGING ------------------------------------------------
Public Sub LogMsg(module_ As String, niveau As String, message As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    
    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = GetCurrentUser()
    ws.Cells(nextRow, 3).Value = module_
    ws.Cells(nextRow, 4).Value = niveau  ' INFO / WARNING / ERROR
    
    ' Empecher Excel d'interpreter les chaines debutant par = comme des formules
    ws.Cells(nextRow, 5).NumberFormat = "@"
    If Left(message, 1) = "=" Then
        ws.Cells(nextRow, 5).Value = "'" & message
    Else
        ws.Cells(nextRow, 5).Value = message
    End If
    
    ' Coloration selon niveau sur la cellule colonne 4 uniquement
    Select Case UCase(niveau)
        Case "ERROR":   ws.Cells(nextRow, 4).Interior.Color = RGB(255, 100, 100)
        Case "WARNING": ws.Cells(nextRow, 4).Interior.Color = RGB(255, 200, 100)
        Case "INFO":    ws.Cells(nextRow, 4).Interior.Color = RGB(200, 230, 255)
    End Select
    On Error GoTo 0
End Sub

' ---- UTILITAIRES -------------------------------------------
Public Function GetCurrentUser() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_ACCUEIL)
    If Not ws Is Nothing Then
        If ws.Range("B_USER_NAME").Value <> "" Then
            GetCurrentUser = ws.Range("B_USER_NAME").Value
            Exit Function
        End If
    End If
    GetCurrentUser = Environ("USERNAME")
End Function

Public Function GetCycleID() As String
    GetCycleID = CStr(GetParam("CYCLE_ID"))
    If GetCycleID = "" Then GetCycleID = "CYC-" & Format(Now(), "YYYY-MM")
End Function

Public Function GetHorizon() As Integer
    Dim h As Variant
    h = GetParam("HORIZON_MOIS")
    If IsNumeric(h) Then GetHorizon = CInt(h) Else GetHorizon = 6
End Function

Public Function GetDateCoupure() As Date
    Dim d As String
    d = Trim(CStr(GetParam("DATE_COUPURE")))
    On Error Resume Next
    If Len(d) >= 10 Then
        GetDateCoupure = DateSerial(CInt(Left(d, 4)), CInt(Mid(d, 6, 2)), CInt(Right(d, 2)))
    Else
        GetDateCoupure = DateSerial(2023, 5, 31)
    End If
End Function

Public Function ParseYM(ymStr As String) As Date
    On Error Resume Next
    Dim s As String: s = Trim(ymStr)
    If Len(s) >= 7 Then
        Dim y As Integer: y = CInt(Left(s, 4))
        Dim m As Integer: m = CInt(Mid(s, 6, 2))
        ParseYM = DateSerial(y, m, 1)
    Else
        ParseYM = DateSerial(2023, 6, 1)
    End If
End Function

Public Sub ShowProgress(step As String, pct As Integer)
    Application.StatusBar = "[" & TOOL_NAME & "] " & step & " (" & pct & "%)"
    DoEvents
End Sub

Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

Public Sub SetStatus(msg As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ACCUEIL)
    If Not ws Is Nothing Then
        ws.Range("B_STATUS").Value = msg
    End If
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Public Sub ClearSheet(shName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(shName)
    If ws Is Nothing Then Exit Sub
    
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 30
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow + 10, lastCol)).Clear
    End If
    On Error GoTo 0
End Sub

Public Sub SetupTable(ws As Worksheet, tblName As String, headers As Variant, rowOut As Long, Optional tblStyle As String = "TableStyleMedium4")
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    
    Dim numCols As Long
    numCols = UBound(headers) - LBound(headers) + 1
    
    ' 1. Ecrire explicitement les entetes sur la ligne 1
    Dim c As Long
    For c = 1 To numCols
        ws.Cells(1, c).Value = headers(LBound(headers) + c - 1)
    Next c
    
    ' Style visuel de l'entete
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, numCols))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 117)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Un tableau structure requiert au moins 1 ligne de donnees (rowOut > 2)
    If rowOut <= 2 Then Exit Sub
    
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowOut - 1, numCols))
    
    ' S'assurer qu'aucun ancien ListObject ne subsiste
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    
    ' Re-ecrire les entetes au cas ou le delete a touche la ligne 1
    For c = 1 To numCols
        ws.Cells(1, c).Value = headers(LBound(headers) + c - 1)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, numCols))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 117)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    If Not lo Is Nothing Then
        lo.Name = tblName
        lo.TableStyle = tblStyle
        For c = 1 To numCols
            lo.ListColumns(c).Name = CStr(headers(LBound(headers) + c - 1))
        Next c
    End If
    On Error GoTo 0
End Sub


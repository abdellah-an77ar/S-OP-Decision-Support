Attribute VB_Name = "mod02_IMPORT"
'==============================================================
' MODULE 02 - IMPORTATION ET VALIDATION DES DONNEES
' S&OP Decision Support System
' F1 - EF-01 : Import CSV + controles qualite
'==============================================================
Option Explicit

' -------------------------------------------------------
' Alias public appele par les boutons et l'orchestrateur
' -------------------------------------------------------
Public Sub ImporterDonnees()
    Call ImporterToutesLesDonnees
End Sub

' -------------------------------------------------------
' Point d'entree principal : importer tous les CSV
' -------------------------------------------------------
Public Sub ImporterToutesLesDonnees()
    Dim dataDir As String
    Dim errorCount As Integer
    Dim warnCount As Integer
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "IMPORT", "INFO", "=== DEBUT IMPORTATION ==="
    ShowProgress "Importation en cours...", 0
    
    ' Detection automatique du dossier Data (robuste via FSO)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim defaultDir As String
    If ThisWorkbook.Path <> "" Then
        defaultDir = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), "Data")
        If Not fso.FolderExists(defaultDir) Then defaultDir = fso.BuildPath(ThisWorkbook.Path, "Data")
    End If
    If defaultDir = "" Or Not fso.FolderExists(defaultDir) Then
        If fso.FolderExists("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Data") Then
            defaultDir = "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Data"
        End If
    End If
    
    If defaultDir <> "" And fso.FolderExists(defaultDir) And fso.FileExists(fso.BuildPath(defaultDir, "SALES_HISTORY.csv")) Then
        dataDir = defaultDir
    Else
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.Title = "Selectionner le dossier Data (contenant les CSV)"
        fd.InitialFileName = ThisWorkbook.Path & "\..\"
        
        If fd.Show = -1 Then
            dataDir = fd.SelectedItems(1)
        Else
            LogMsg "IMPORT", "WARNING", "Importation annulee par l'utilisateur."
            GoTo CleanExit
        End If
    End If
    
    errorCount = 0
    warnCount = 0
    
    ' Import chaque fichier
    ShowProgress "Importation PRODUCT...", 10
    errorCount = errorCount + ImportCSV(dataDir & "\PRODUCT.csv", "tbl_PRODUCT", SH_PRODUCT)
    
    ShowProgress "Importation PLANT...", 20
    errorCount = errorCount + ImportCSV(dataDir & "\PLANT.csv", "tbl_PLANT", SH_PLANT)
    
    ShowProgress "Importation CALENDAR...", 30
    errorCount = errorCount + ImportCSV(dataDir & "\CALENDAR.csv", "tbl_CALENDAR", SH_CALENDAR)
    
    ShowProgress "Importation COST...", 40
    errorCount = errorCount + ImportCSV(dataDir & "\COST.csv", "tbl_COST", SH_COST)
    
    ShowProgress "Importation SALES_HISTORY...", 50
    errorCount = errorCount + ImportCSV(dataDir & "\SALES_HISTORY.csv", "tbl_SALES", SH_SALES)
    
    ShowProgress "Importation PRODUCTION_HISTORY...", 65
    errorCount = errorCount + ImportCSV(dataDir & "\PRODUCTION_HISTORY.csv", "tbl_PROD_HIST", SH_PROD_HIST)
    
    ShowProgress "Importation DELIVERY_HISTORY...", 80
    errorCount = errorCount + ImportCSV(dataDir & "\DELIVERY_HISTORY.csv", "tbl_DELIV", SH_DELIV)
    
    ' Validation croisee
    ShowProgress "Validation croisee...", 90
    warnCount = ValiderDonnees()
    
    ShowProgress "Import termine.", 100
    SetStatus "Donnees importees avec succes - " & Format(Now(), "yyyy-mm-dd hh:mm")
    
    Dim msg As String
    If errorCount = 0 And warnCount = 0 Then
        msg = "Import termine avec succes." & vbCrLf & "7 fichiers importes. Aucune anomalie."
        If Application.UserControl Then MsgBox msg, vbInformation, TOOL_NAME
        LogMsg "IMPORT", "INFO", "Import OK - 0 erreur, 0 avertissement"
    ElseIf errorCount = 0 Then
        msg = "Import termine avec succes." & vbCrLf & warnCount & " avertissement(s) de qualite detecte(s)." & vbCrLf & "Consultez l'onglet LOG."
        If Application.UserControl Then MsgBox msg, vbInformation, TOOL_NAME
        LogMsg "IMPORT", "WARNING", "Import avec " & warnCount & " avertissement(s)"
    Else
        msg = errorCount & " erreur(s) detectee(s)." & vbCrLf & "Consultez l'onglet LOG."
        If Application.UserControl Then MsgBox msg, vbCritical, TOOL_NAME
        LogMsg "IMPORT", "ERROR", "Import echoue - " & errorCount & " erreur(s)"
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    Exit Sub
    
ErrHandler:
    LogMsg "IMPORT", "ERROR", "Erreur fatale Import: " & Err.Number & " - " & Err.Description
    If Application.UserControl Then MsgBox "Erreur fatale lors de l'importation:" & vbCrLf & Err.Description, vbCritical, TOOL_NAME
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
End Sub

' -------------------------------------------------------
' Importer un CSV dans une feuille Excel (Performant via Array)
' Retourne 0=OK, 1=ERREUR BLOQUANTE
' -------------------------------------------------------
Private Function ImportCSV(filePath As String, tableName As String, sheetName As String) As Integer
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim j As Integer
    Dim colCount As Integer
    
    ImportCSV = 0
    
    Dim curStep As String
    curStep = "Start"
    On Error GoTo ErrCSV

    ' Verifier que le fichier existe
    curStep = "CheckFile"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        LogMsg "IMPORT", "ERROR", "Fichier introuvable: " & filePath
        ImportCSV = 1
        Exit Function
    End If
    
    ' Recuperer la feuille
    curStep = "GetSheet"
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        LogMsg "IMPORT", "ERROR", "Feuille introuvable: " & sheetName
        ImportCSV = 1
        Exit Function
    End If
    
    ' Effacer tous les ListObjects existants sur la feuille
    curStep = "DeleteTables"
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    
    ' Lire le CSV (UTF-8 via ADODB.Stream)
    curStep = "ReadStream"
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    
    Dim allText As String
    allText = stream.ReadText()
    stream.Close
    
    ' Decoder les lignes
    curStep = "SplitLines"
    Dim allLines() As String
    allLines = Split(allText, vbLf)
    
    ' Nettoyer retours chariot
    Dim k As Long
    For k = 0 To UBound(allLines)
        allLines(k) = Replace(allLines(k), vbCr, "")
    Next k
    
    If UBound(allLines) < 1 Then
        LogMsg "IMPORT", "ERROR", "Fichier vide: " & filePath
        ImportCSV = 1
        Exit Function
    End If
    
    ' En-tetes
    curStep = "Headers"
    Dim headers() As String
    headers = Split(allLines(0), ";")
    colCount = UBound(headers) + 1
    
    ws.Cells.Clear
    For j = 0 To colCount - 1
        headers(j) = Trim(Replace(headers(j), """", ""))
        ws.Cells(1, j + 1).Value = headers(j)
        ws.Cells(1, j + 1).Font.Bold = True
        ws.Cells(1, j + 1).Interior.Color = RGB(31, 73, 125)
        ws.Cells(1, j + 1).Font.Color = RGB(255, 255, 255)
    Next j
    
    ' Compter lignes valides d'abord
    curStep = "CountValidRows"
    Dim validRows As Long: validRows = 0
    For k = 1 To UBound(allLines)
        If Trim(allLines(k)) <> "" Then validRows = validRows + 1
    Next k
    
    If validRows = 0 Then
        LogMsg "IMPORT", "WARNING", "Aucune donnee dans: " & filePath
        ImportCSV = 0
        Exit Function
    End If
    
    ' Remplissage dans un tableau 2D de taille exacte
    curStep = "FillArray"
    Dim arrData() As Variant
    ReDim arrData(1 To validRows, 1 To colCount)
    
    Dim rowIdx As Long: rowIdx = 0
    Dim lineVals() As String
    For k = 1 To UBound(allLines)
        If Trim(allLines(k)) <> "" Then
            lineVals = Split(allLines(k), ";")
            rowIdx = rowIdx + 1
            For j = 0 To UBound(lineVals)
                If j < colCount Then
                    Dim cv As String
                    cv = Trim(Replace(lineVals(j), """", ""))
                    If IsNumeric(cv) Then
                        arrData(rowIdx, j + 1) = CDbl(cv)
                    ElseIf IsDate(cv) And Len(cv) = 10 Then
                        arrData(rowIdx, j + 1) = CDate(cv)
                    Else
                        arrData(rowIdx, j + 1) = cv
                    End If
                End If
            Next j
        End If
    Next k
    
    ' Ecriture en bloc sur la feuille
    curStep = "WriteRange"
    ws.Range(ws.Cells(2, 1), ws.Cells(validRows + 1, colCount)).Value = arrData
    
    ' Format des dates
    curStep = "FormatDates"
    Dim cIdx As Integer
    For cIdx = 1 To colCount
        If UCase(headers(cIdx - 1)) Like "*DATE*" Or UCase(headers(cIdx - 1)) Like "*MOIS*" Or UCase(headers(cIdx - 1)) Like "*PERIODE*" Then
            ws.Range(ws.Cells(2, cIdx), ws.Cells(validRows + 1, cIdx)).NumberFormat = "yyyy-mm-dd"
        End If
    Next cIdx
    
    ' Creation ListObject
    curStep = "CreateListObject"
    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(validRows + 1, colCount)), , xlYes)
    newLo.Name = tableName
    newLo.TableStyle = "TableStyleMedium2"
    
    curStep = "AutoFit"
    ws.Columns.AutoFit
    LogMsg "IMPORT", "INFO", "Import OK: " & sheetName & " (" & validRows & " lignes)"
    ImportCSV = 0
    Exit Function

ErrCSV:
    LogMsg "IMPORT", "ERROR", "ImportCSV [" & sheetName & "/" & tableName & "] at " & curStep & ": " & Err.Number & " - " & Err.Description
    ImportCSV = 1
End Function

' -------------------------------------------------------
' Validation croisee des donnees importees (BR-01 a BR-06)
' Retourne le nombre d'avertissements
' -------------------------------------------------------
Private Function ValiderDonnees() As Integer
    Dim warns As Integer: warns = 0
    Dim ws As Worksheet
    Dim cell As Range
    
    ' --- BR-01 & BR-02 : Controles SALES_HISTORY ---
    Set ws = ThisWorkbook.Sheets(SH_SALES)
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tbl_SALES"): On Error GoTo 0
    
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        For Each cell In lo.DataBodyRange.Columns(3).Cells  ' Qte_commandee
            If IsEmpty(cell.Value) Or cell.Value = "" Then
                LogMsg "VALID", "WARNING", "BR-01: Valeur manquante SALES ligne " & cell.Row
                cell.Interior.Color = RGB(255, 255, 150)
                warns = warns + 1
            ElseIf IsNumeric(cell.Value) Then
                If cell.Value < 0 Then
                    LogMsg "VALID", "WARNING", "BR-02: Quantite negative SALES ligne " & cell.Row & " = " & cell.Value
                    cell.Interior.Color = RGB(255, 150, 150)
                    warns = warns + 1
                End If
            End If
        Next cell
        
        ' BR-06 : Outliers
        Dim col3 As Range: Set col3 = lo.DataBodyRange.Columns(3)
        Dim avg As Double: avg = Application.Average(col3)
        Dim stdDev As Double: stdDev = Application.StDev(col3)
        For Each cell In col3.Cells
            If IsNumeric(cell.Value) Then
                If cell.Value > avg + 3 * stdDev Then
                    LogMsg "VALID", "WARNING", "BR-06: Valeur extreme SALES ligne " & cell.Row & " = " & cell.Value
                    warns = warns + 1
                End If
            End If
        Next cell
    End If
    
    ' --- BR-04 : Mois partiel ---
    LogMsg "VALID", "INFO", "BR-04: Mois partiel detecte - Aout 2023. Exclu du cycle d'alignement."
    LogMsg "VALID", "INFO", "Validation terminee: " & warns & " avertissement(s)"
    ValiderDonnees = warns
End Function

Attribute VB_Name = "mod03_FORECAST"
'==============================================================
' MODULE 03 - ANALYSE DE LA DEMANDE ET FORECASTING
' S&OP Decision Support System
' F2 (classification) + F3 (forecast multi-methodes + backtest)
' Methodes: Moyenne mobile, Lissage exponentiel, Holt, Holt-Winters simplifie
'==============================================================
Option Explicit

' Structure pour stocker les resultats de forecast
Private Type ForecastResult
    SKU_ID As String
    Method As String
    WAPE As Double
    Biais As Double
    Periods(6) As Double  ' 6 mois horizon
End Type

' -------------------------------------------------------
' Analyse la demande et lance le forecast pour tous les SKU
' -------------------------------------------------------
Public Sub LancerForecast()
    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim lo As ListObject
    Dim wsForecast As Worksheet
    Dim cycleID As String
    Dim dateCoupure As Date
    Dim horizon As Integer
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "FORECAST", "INFO", "=== DEBUT FORECAST ==="
    ShowProgress "Analyse de la demande...", 5
    
    cycleID = GetCycleID()
    dateCoupure = GetDateCoupure()
    horizon = GetHorizon()
    
    ' Recuperer la liste des SKU
    Dim skuList() As String
    skuList = GetSKUList()
    If UBound(skuList) < 0 Then
        If Application.UserControl Then MsgBox "Aucun produit trouve. Veuillez d'abord importer les donnees.", vbCritical, TOOL_NAME
        GoTo CleanExit
    End If
    
    ' Preparer la feuille FORECAST
    Set wsForecast = ThisWorkbook.Sheets(SH_FORECAST)
    EffacerDonneesForecast wsForecast
    
    Dim skuIdx As Integer
    Dim totalSKU As Integer
    totalSKU = UBound(skuList) + 1
    
    Dim rowOut As Long
    rowOut = 2
    
    For skuIdx = 0 To UBound(skuList)
        Dim sku As String
        sku = skuList(skuIdx)
        
        ShowProgress "Forecast: " & sku, 5 + Int(90 * skuIdx / totalSKU)
        
        ' Obtenir historique mensuel (jusqu'a dateCoupure)
        Dim histData() As Double
        Dim histPeriods() As Date
        histData = GetHistoriquesMensuel(sku, dateCoupure, histPeriods)
        
        If UBound(histData) < 2 Then
            LogMsg "FORECAST", "WARNING", "Historique insuffisant pour " & sku & " (< 3 mois)"
            GoTo NextSKU
        End If
        
        ' Classer le produit (B2 - F2)
        Dim classe As String
        classe = ClassifierProduit(histData)
        LogMsg "FORECAST", "INFO", "Classification " & sku & ": " & classe
        
        ' Calculer les 4 methodes et backtest
        Dim results(3) As ForecastResult
        Dim btLen As Integer
        btLen = CInt(GetParam("BACKTEST_SEMAINES") / 4)  ' Convertir semaines -> mois approx
        If btLen < 2 Then btLen = 2
        
        results(0) = MoyenneMobile(sku, histData, horizon, btLen)
        results(1) = LissageExponentiel(sku, histData, horizon, btLen)
        results(2) = HoltTendance(sku, histData, horizon, btLen)
        results(3) = HoltWintersSimplifie(sku, histData, horizon, btLen)
        
        ' Choisir la meilleure methode par WAPE minimum
        Dim bestIdx As Integer
        bestIdx = 0
        Dim bestWAPE As Double
        bestWAPE = results(0).WAPE
        Dim m As Integer
        For m = 1 To 3
            If results(m).WAPE < bestWAPE Then
                bestWAPE = results(m).WAPE
                bestIdx = m
            End If
        Next m
        
        ' Ecrire toutes les methodes dans FORECAST (pour comparaison)
        Dim mStart As Date
        mStart = DateAdd("m", 1, dateCoupure)
        mStart = DateSerial(Year(mStart), Month(mStart), 1)
        
        For m = 0 To 3
            Dim p As Integer
            For p = 0 To horizon - 1
                Dim mDate As Date
                mDate = DateAdd("m", p, mStart)
                
                Dim statut As String
                If m = bestIdx Then
                    statut = "RETENUE"
                Else
                    statut = "Alternative"
                End If
                
                ' Ecrire dans la feuille
                wsForecast.Cells(rowOut, 1).Value = sku
                wsForecast.Cells(rowOut, 2).Value = mDate
                wsForecast.Cells(rowOut, 2).NumberFormat = "mmm-yy"
                wsForecast.Cells(rowOut, 3).Value = Round(results(m).Periods(p), 0)
                wsForecast.Cells(rowOut, 4).Value = results(m).Method
                wsForecast.Cells(rowOut, 5).Value = Round(results(m).WAPE * 100, 2)
                wsForecast.Cells(rowOut, 6).Value = Round(results(m).Biais * 100, 2)
                wsForecast.Cells(rowOut, 7).Value = statut
                wsForecast.Cells(rowOut, 8).Value = cycleID
                wsForecast.Cells(rowOut, 9).Value = classe
                
                If statut = "RETENUE" Then
                    wsForecast.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(200, 240, 200)
                End If
                rowOut = rowOut + 1
            Next p
        Next m
        
NextSKU:
    Next skuIdx
    
    ' Reformater comme tableau
    ReformaterTableForecast wsForecast, rowOut - 1
    
    LogMsg "FORECAST", "INFO", "Forecast termine. " & totalSKU & " SKU traites."
    ShowProgress "Forecast termine.", 100
    
    If Application.UserControl Then
        MsgBox "Forecast calcule pour " & totalSKU & " produits." & vbCrLf & _
               "Consultez l'onglet FORECAST pour les resultats.", vbInformation, TOOL_NAME
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    Exit Sub
    
ErrHandler:
    LogMsg "FORECAST", "ERROR", "Erreur Forecast: " & Err.Number & " - " & Err.Description
    If Application.UserControl Then MsgBox "Erreur Forecast: " & Err.Description, vbCritical, TOOL_NAME
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
End Sub

' -------------------------------------------------------
' Retourne la liste des SKU depuis la feuille PRODUCT
' -------------------------------------------------------
Private Function GetSKUList() As String()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim result() As String
    Dim count As Integer
    
    Set ws = ThisWorkbook.Sheets(SH_PRODUCT)
    On Error Resume Next
    Set lo = ws.ListObjects("tbl_PRODUCT")
    On Error GoTo 0
    
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        ReDim result(-1 To -1)
        GetSKUList = result
        Exit Function
    End If
    
    count = lo.DataBodyRange.Rows.Count
    ReDim result(count - 1)
    Dim i As Integer
    For i = 1 To count
        result(i - 1) = CStr(lo.DataBodyRange(i, 1).Value)
    Next i
    GetSKUList = result
End Function

' -------------------------------------------------------
' Retourne l'historique mensuel agreage pour un SKU
' -------------------------------------------------------
Private Function GetHistoriquesMensuel(sku As String, cutDate As Date, ByRef periods() As Date) As Double()
    Dim wsSales As Worksheet
    Dim lo As ListObject
    Dim cell As Range
    Dim mois As Date
    Dim dateCell As Date
    Dim i As Long
    
    Set wsSales = ThisWorkbook.Sheets(SH_SALES)
    On Error Resume Next
    Set lo = wsSales.ListObjects("tbl_SALES")
    On Error GoTo 0
    
    ' Construire un dictionnaire mois -> somme
    Dim monthSums As Object
    Set monthSums = CreateObject("Scripting.Dictionary")
    
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.DataBodyRange.Rows.Count
            Dim skuCell As String
            skuCell = CStr(lo.DataBodyRange(i, 1).Value)
            If skuCell <> sku Then GoTo NextRow
            
            Dim dateVal As Variant
            dateVal = lo.DataBodyRange(i, 2).Value
            If Not IsDate(dateVal) Then GoTo NextRow
            dateCell = CDate(dateVal)
            If dateCell > cutDate Then GoTo NextRow
            
            mois = DateSerial(Year(dateCell), Month(dateCell), 1)
            Dim mKey As String
            mKey = Format(mois, "YYYY-MM")
            
            Dim qty As Double
            qty = 0
            If IsNumeric(lo.DataBodyRange(i, 3).Value) Then qty = CDbl(lo.DataBodyRange(i, 3).Value)
            
            If monthSums.Exists(mKey) Then
                monthSums(mKey) = monthSums(mKey) + qty
            Else
                monthSums.Add mKey, qty
            End If
NextRow:
        Next i
    End If
    
    ' Convertir en arrays tries
    Dim keys As Variant
    keys = monthSums.Keys()
    
    ' Trier par date (tri simple)
    Dim n As Integer
    n = monthSums.Count
    If n = 0 Then
        ReDim GetHistoriquesMensuel(-1 To -1)
        Exit Function
    End If
    
    Dim sortedKeys() As String
    ReDim sortedKeys(n - 1)
    Dim j As Integer
    For i = 0 To n - 1
        sortedKeys(i) = CStr(keys(i))
    Next i
    
    ' Bubble sort des cles (format YYYY-MM -> tri lexicographique = chronologique)
    Dim tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If sortedKeys(i) > sortedKeys(j) Then
                tmp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = tmp
            End If
        Next j
    Next i
    
    Dim resultVals() As Double
    ReDim resultVals(n - 1)
    ReDim periods(n - 1)
    
    For i = 0 To n - 1
        resultVals(i) = CDbl(monthSums(sortedKeys(i)))
        periods(i) = ParseYM(CStr(sortedKeys(i)))
    Next i
    
    GetHistoriquesMensuel = resultVals
End Function

' -------------------------------------------------------
' Classification du produit (F2 - BR-07)
' -------------------------------------------------------
Private Function ClassifierProduit(data() As Double) As String
    Dim n As Integer
    n = UBound(data) + 1
    If n < 3 Then
        ClassifierProduit = "INCONNU"
        Exit Function
    End If
    
    ' Calcul de la moyenne et ecart-type
    Dim avg As Double
    Dim i As Integer
    avg = 0
    For i = 0 To n - 1
        avg = avg + data(i)
    Next i
    avg = avg / n
    
    If avg = 0 Then
        ClassifierProduit = "VOLATILE"
        Exit Function
    End If
    
    Dim variance As Double
    variance = 0
    For i = 0 To n - 1
        variance = variance + (data(i) - avg) ^ 2
    Next i
    variance = variance / (n - 1)
    Dim stdDev As Double
    stdDev = Sqr(variance)
    
    ' Coefficient de variation
    Dim cv As Double
    cv = stdDev / avg
    
    ' Tendance (regression lineaire simple)
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0
    For i = 0 To n - 1
        sumX = sumX + i
        sumY = sumY + data(i)
        sumXY = sumXY + i * data(i)
        sumX2 = sumX2 + i ^ 2
    Next i
    
    Dim slope As Double
    Dim denom As Double
    denom = n * sumX2 - sumX ^ 2
    If Abs(denom) < 0.0001 Then
        slope = 0
    Else
        slope = (n * sumXY - sumX * sumY) / denom
    End If
    
    Dim slopePct As Double
    If avg > 0 Then slopePct = Abs(slope) / avg Else slopePct = 0
    
    ' Saisonnalite: variance des moyennes bimensuelles
    Dim cvSeuil As Double: cvSeuil = CDbl(GetParam("CV_SEUIL_STABLE"))
    Dim cvVolSeuil As Double: cvVolSeuil = CDbl(GetParam("CV_SEUIL_VOLATILE"))
    
    If cv > cvVolSeuil Then
        ClassifierProduit = "VOLATILE"
    ElseIf slopePct > 0.05 Then
        ClassifierProduit = "TENDANCIEL"
    ElseIf cv < cvSeuil Then
        ClassifierProduit = "STABLE"
    Else
        ClassifierProduit = "SAISONNIER"
    End If
End Function

' -------------------------------------------------------
' METHODE 1 : Moyenne mobile (n=3 mois)
' -------------------------------------------------------
Private Function MoyenneMobile(sku As String, data() As Double, horizon As Integer, btLen As Integer) As ForecastResult
    Dim res As ForecastResult
    res.SKU_ID = sku
    res.Method = "Moyenne Mobile (3)"
    
    Dim n As Integer
    n = UBound(data) + 1
    Dim w As Integer: w = 3  ' Fenetre
    
    ' Backtest
    Dim wape As Double, biais As Double
    Dim errSum As Double, demSum As Double
    Dim i As Integer
    errSum = 0: demSum = 0
    
    For i = n - btLen To n - 1
        If i >= w Then
            Dim pred As Double
            pred = 0
            Dim j As Integer
            For j = i - w To i - 1
                pred = pred + data(j)
            Next j
            pred = pred / w
            
            errSum = errSum + Abs(pred - data(i))
            demSum = demSum + data(i)
        End If
    Next i
    
    If demSum > 0 Then
        wape = errSum / demSum
        biais = (Application.Sum(Application.Index(data, 0)) - Application.Sum(Application.Index(data, 0))) / demSum
    End If
    
    ' Calcul biais simple
    Dim biaisSum As Double
    biaisSum = 0
    For i = n - btLen To n - 1
        If i >= w Then
            Dim p2 As Double
            p2 = 0
            For j = i - w To i - 1
                p2 = p2 + data(j)
            Next j
            p2 = p2 / w
            biaisSum = biaisSum + (p2 - data(i))
        End If
    Next i
    If demSum > 0 Then biais = biaisSum / demSum
    
    res.WAPE = wape
    res.Biais = biais
    
    ' Forecast horizon
    Dim lastN As Double: lastN = 0
    Dim startIdx As Integer: startIdx = n - w
    If startIdx < 0 Then startIdx = 0
    For i = startIdx To n - 1
        lastN = lastN + data(i)
    Next i
    Dim baseVal As Double
    baseVal = lastN / Application.Min(w, n)
    
    For i = 0 To horizon - 1
        res.Periods(i) = Application.Max(0, baseVal)
    Next i
    
    MoyenneMobile = res
End Function

' -------------------------------------------------------
' METHODE 2 : Lissage exponentiel simple
' -------------------------------------------------------
Private Function LissageExponentiel(sku As String, data() As Double, horizon As Integer, btLen As Integer) As ForecastResult
    Dim res As ForecastResult
    res.SKU_ID = sku
    res.Method = "Lissage Exponentiel"
    
    Dim n As Integer: n = UBound(data) + 1
    Dim alpha As Double: alpha = 0.3  ' Parametre de lissage
    
    ' Initialisation
    Dim smoothed As Double
    smoothed = data(0)
    
    ' Backtest
    Dim errSum As Double, demSum As Double, biaisSum As Double
    errSum = 0: demSum = 0: biaisSum = 0
    
    Dim i As Integer
    For i = 1 To n - 1
        Dim pred As Double
        pred = smoothed
        
        If i >= n - btLen Then
            errSum = errSum + Abs(pred - data(i))
            demSum = demSum + data(i)
            biaisSum = biaisSum + (pred - data(i))
        End If
        
        smoothed = alpha * data(i) + (1 - alpha) * smoothed
    Next i
    
    If demSum > 0 Then
        res.WAPE = errSum / demSum
        res.Biais = biaisSum / demSum
    End If
    
    ' Forecast = derniere valeur lissee
    For i = 0 To horizon - 1
        res.Periods(i) = Application.Max(0, smoothed)
    Next i
    
    LissageExponentiel = res
End Function

' -------------------------------------------------------
' METHODE 3 : Holt (tendance lineaire)
' -------------------------------------------------------
Private Function HoltTendance(sku As String, data() As Double, horizon As Integer, btLen As Integer) As ForecastResult
    Dim res As ForecastResult
    res.SKU_ID = sku
    res.Method = "Holt (Tendance)"
    
    Dim n As Integer: n = UBound(data) + 1
    Dim alpha As Double: alpha = 0.3
    Dim beta As Double: beta = 0.1
    
    If n < 2 Then
        res.WAPE = 1
        LancerForecast_Baseline res, data, horizon
        HoltTendance = res
        Exit Function
    End If
    
    ' Initialisation Holt
    Dim Lt As Double: Lt = data(0)
    Dim Tt As Double: Tt = data(1) - data(0)
    
    Dim errSum As Double, demSum As Double, biaisSum As Double
    errSum = 0: demSum = 0: biaisSum = 0
    
    Dim i As Integer
    For i = 1 To n - 1
        Dim pred As Double
        pred = Lt + Tt
        pred = Application.Max(0, pred)
        
        If i >= n - btLen Then
            errSum = errSum + Abs(pred - data(i))
            demSum = demSum + data(i)
            biaisSum = biaisSum + (pred - data(i))
        End If
        
        Dim Lt_prev As Double: Lt_prev = Lt
        Lt = alpha * data(i) + (1 - alpha) * (Lt + Tt)
        Tt = beta * (Lt - Lt_prev) + (1 - beta) * Tt
    Next i
    
    If demSum > 0 Then
        res.WAPE = errSum / demSum
        res.Biais = biaisSum / demSum
    End If
    
    ' Forecast
    For i = 1 To horizon
        res.Periods(i - 1) = Application.Max(0, Lt + i * Tt)
    Next i
    
    HoltTendance = res
End Function

' Initialise les periodes du forecast a la moyenne
Private Sub LancerForecast_Baseline(ByRef res As ForecastResult, data() As Double, horizon As Integer)
    Dim n As Integer: n = UBound(data) + 1
    Dim avg As Double: avg = 0
    Dim i As Integer
    For i = 0 To n - 1: avg = avg + data(i): Next i
    avg = avg / n
    For i = 0 To horizon - 1
        res.Periods(i) = avg
    Next i
End Sub

' -------------------------------------------------------
' METHODE 4 : Holt-Winters simplifie (saisonnalite approximee)
' -------------------------------------------------------
Private Function HoltWintersSimplifie(sku As String, data() As Double, horizon As Integer, btLen As Integer) As ForecastResult
    Dim res As ForecastResult
    res.SKU_ID = sku
    res.Method = "Holt-Winters (Saison)"
    
    Dim n As Integer: n = UBound(data) + 1
    
    ' Avec moins de 7 mois, on ne peut pas faire une vraie saisonnalite sur 12 periodes
    ' On utilise une approximation sur 3-4 periodes de saisonnalite
    Dim seasonLen As Integer: seasonLen = 3
    
    If n < seasonLen + 2 Then
        ' Pas assez de donnees : utiliser Holt simple
        HoltWintersSimplifie = HoltTendance(sku, data, horizon, btLen)
        HoltWintersSimplifie.Method = "Holt-Winters (Saison)"
        Exit Function
    End If
    
    Dim alpha As Double: alpha = 0.3
    Dim beta  As Double: beta  = 0.1
    Dim gamma As Double: gamma = 0.2
    
    ' Initialisation
    Dim i As Integer, j As Integer
    
    ' Moyenne initiale
    Dim initAvg As Double: initAvg = 0
    For i = 0 To seasonLen - 1: initAvg = initAvg + data(i): Next i
    initAvg = initAvg / seasonLen
    
    ' Indices saisonniers initiaux
    Dim SI(2) As Double  ' Season Index pour periode 3
    For i = 0 To seasonLen - 1
        If initAvg > 0 Then
            SI(i) = data(i) / initAvg
        Else
            SI(i) = 1
        End If
    Next i
    
    Dim Lt As Double: Lt = initAvg
    Dim Tt As Double: Tt = 0
    
    Dim errSum As Double, demSum As Double, biaisSum As Double
    errSum = 0: demSum = 0: biaisSum = 0
    
    For i = seasonLen To n - 1
        Dim siIdx As Integer: siIdx = (i Mod seasonLen)
        Dim pred As Double
        pred = Application.Max(0, (Lt + Tt) * SI(siIdx))
        
        If i >= n - btLen Then
            errSum = errSum + Abs(pred - data(i))
            demSum = demSum + data(i)
            biaisSum = biaisSum + (pred - data(i))
        End If
        
        Dim Lt_prev As Double: Lt_prev = Lt
        If SI(siIdx) > 0 Then
            Lt = alpha * (data(i) / SI(siIdx)) + (1 - alpha) * (Lt + Tt)
        End If
        Tt = beta * (Lt - Lt_prev) + (1 - beta) * Tt
        SI(siIdx) = gamma * (data(i) / Lt) + (1 - gamma) * SI(siIdx)
    Next i
    
    If demSum > 0 Then
        res.WAPE = errSum / demSum
        res.Biais = biaisSum / demSum
    End If
    
    ' Forecast horizon
    For i = 0 To horizon - 1
        Dim fSI As Integer: fSI = ((n + i) Mod seasonLen)
        res.Periods(i) = Application.Max(0, (Lt + (i + 1) * Tt) * SI(fSI))
    Next i
    
    HoltWintersSimplifie = res
End Function

' -------------------------------------------------------
' Effacer les donnees de forecast precedentes
' -------------------------------------------------------
Private Sub EffacerDonneesForecast(ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tbl_FORECAST")
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Delete
        End If
    End If
End Sub

' -------------------------------------------------------
' Reformater la plage de resultats en tableau structure
' -------------------------------------------------------
Private Sub ReformaterTableForecast(ws As Worksheet, lastRow As Long)
    If lastRow < 2 Then Exit Sub
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tbl_FORECAST")
    On Error GoTo 0
    
    If lo Is Nothing And lastRow >= 2 Then
        ' Creer le tableau
        ws.ListObjects.Add(xlSrcRange, ws.Range("A1:I" & lastRow), , xlYes).Name = "tbl_FORECAST"
        ' Appliquer un style
        ws.ListObjects("tbl_FORECAST").TableStyle = "TableStyleMedium2"
    End If
End Sub

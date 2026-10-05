$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open("c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\AtlasFood_SOP.xlsm")
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.CodeModule.AddFromString(@"
Public Sub TestExactLoop()
    On Error GoTo ErrH
    Dim fnDbg As Integer: fnDbg = FreeFile
    Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_err.txt" For Output As #fnDbg
    Print #fnDbg, "Entered TestExactLoop"
    Close #fnDbg
    Dim wsDp As Worksheet:    Set wsDp    = ThisWorkbook.Sheets(SH_DEMAND_PLAN)
    Dim wsCap As Worksheet:   Set wsCap   = ThisWorkbook.Sheets(SH_CAPACITY)
    Dim wsPlant As Worksheet: Set wsPlant = ThisWorkbook.Sheets(SH_PLANT)
    Dim wsProd As Worksheet:  Set wsProd  = ThisWorkbook.Sheets(SH_PRODUCT)

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

    Dim arrKeys As Variant
    arrKeys = dictCharge.Keys
    Dim ik As Long
    Dim rowOut As Long: rowOut = 2
    Dim fn As Integer: fn = FreeFile
    Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_loop.txt" For Output As #fn
    Print #fn, "arrKeys count: " & (UBound(arrKeys) + 1)
    Close #fn
    For ik = 0 To UBound(arrKeys)
        Dim k As String: k = CStr(arrKeys(ik))
        fn = FreeFile: Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_loop.txt" For Append As #fn
        Print #fn, "ik=" & ik & " k=" & k
        Close #fn
        Dim kp() As String: kp = Split(CStr(k), "|")
        Dim kPlant As String: kPlant = kp(0)
        Dim kMois As Date: kMois = ParseYM(kp(1))
        Dim charge As Double: charge = CDbl(dictCharge(k))
        Debug.Print "  charge=" & charge
        Dim capN As Double: Dim capHS As Double: Dim capST As Double
        If dictCap.Exists(kPlant) Then
            Dim cParts() As String: cParts = Split(CStr(dictCap(kPlant)), "|")
            capN = CDbl(cParts(0)): capHS = CDbl(cParts(1)): capST = CDbl(cParts(2))
        Else: capN = 5000: capHS = 1000: capST = 500
        End If
        Debug.Print "  capN=" & capN
        Dim util As Double: If capN > 0 Then util = charge / capN Else util = 0
        Dim st As String
        If util > 1 Then st = "SURCHARGE" ElseIf util > seuil Then st = "ALERTE" Else st = "OK"
        Debug.Print "  st=" & st
        wsCap.Cells(rowOut, 1).Value = kPlant
        wsCap.Cells(rowOut, 2).Value = kMois: wsCap.Cells(rowOut, 2).NumberFormat = "mmm-yy"
        wsCap.Cells(rowOut, 3).Value = Round(capN, 0)
        wsCap.Cells(rowOut, 4).Value = Round(capHS, 0)
        wsCap.Cells(rowOut, 5).Value = Round(capST, 0)
        wsCap.Cells(rowOut, 6).Value = Round(charge, 0)
        wsCap.Cells(rowOut, 7).Value = Round(util * 100, 1)
        wsCap.Cells(rowOut, 8).Value = st
        wsCap.Cells(rowOut, 9).Value = cycleID
        Debug.Print "  cells written"
        ' Select Case st
        '     Case "SURCHARGE": wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(255, 150, 150)
        '     Case "ALERTE":    wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(255, 230, 150)
        '     Case Else:        wsCap.Range("A" & rowOut & ":I" & rowOut).Interior.Color = RGB(200, 240, 200)
        ' End Select
        Debug.Print "  color skipped"
        rowOut = rowOut + 1
    Next ik
    Debug.Print "Loop finished, rowOut=" & rowOut

    If rowOut > 2 Then
        Dim lo As ListObject
        On Error Resume Next: Set lo = wsCap.ListObjects("tbl_CAPACITY"): On Error GoTo 0
        If lo Is Nothing Then
            Set lo = wsCap.ListObjects.Add(xlSrcRange, wsCap.Range("A1:I" & (rowOut - 1)), , xlYes)
            lo.Name = "tbl_CAPACITY": lo.TableStyle = "TableStyleMedium5"
        End If
    End If
    Debug.Print "TestExactLoop COMPLETE!"
    Exit Sub
ErrH:
    Dim fnErr As Integer: fnErr = FreeFile
    Open "c:\Users\RZR\Documents\EMINES\CI2A\app vba\Application\dbg_err.txt" For Append As #fnErr
    Print #fnErr, "ERROR: " & Err.Number & " - " & Err.Description
    Close #fnErr
End Sub
"@)

    Write-Host "Calling TestExactLoop..."
    $xl.Run("TestExactLoop")
    Write-Host "TestExactLoop PASS!"

} catch {
    Write-Host "Error: $_"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}

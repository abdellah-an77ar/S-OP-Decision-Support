Option Explicit
'==============================================================
' ThisWorkbook - Evenements niveau classeur
' AtlasFood S&OP DSS
'==============================================================

Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Sheets("ACCUEIL").Activate
    Application.StatusBar = "AtlasFood S&OP DSS v1.0 - Pret"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

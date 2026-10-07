VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_AjustementCommercial 
   Caption         =   "Ajustement Commercial de la Demande (T5)"
   ClientHeight    =   6240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "frm_AjustementCommercial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_AjustementCommercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    InitialiserCaptions
    InitialiserFormulaire
End Sub

Private Sub InitialiserCaptions()
    Me.Caption = "Arbitrage Commercial S&OP (T5)"
    lbl_Title.Caption = "  AJUSTEMENT COMMERCIAL DU DEMAND PLAN"
    lbl_SKU.Caption = "Produit (SKU) :"
    lbl_Mois.Caption = "Mois (Horizon) :"
    lbl_PrevStat.Caption = "Pr" & Chr(233) & "vision Stat :"
    lbl_ValPrevStat.Caption = "0 unit" & Chr(233) & "s"
    lbl_TypeAjust.Caption = "Type Ajustement :"
    lbl_ValeurAjust.Caption = "Valeur :"
    lbl_DemandeFinale.Caption = "Demande Valid" & Chr(233) & "e :"
    lbl_ValDemandeFinale.Caption = "0 unit" & Chr(233) & "s"
    lbl_Motif.Caption = "Motif (OBLIGATOIRE) :"
    btn_Enregistrer.Caption = "Valider Ajustement"
    btn_ReinitialiserAjust.Caption = "Remettre " & Chr(224) & " Z" & Chr(233) & "ro"
    btn_Fermer.Caption = "Fermer"
End Sub

Private Sub InitialiserFormulaire()
    cbo_Type_Ajustement.Clear
    cbo_Type_Ajustement.AddItem "Pourcentage (%)"
    cbo_Type_Ajustement.AddItem "Volume Absolu (+/-)"
    cbo_Type_Ajustement.ListIndex = 0
    
    cbo_SKU_ID.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.sheets(SH_PRODUCT)
    Dim lo As ListObject: Set lo = ws.ListObjects("tbl_PRODUCT")
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.count
            cbo_SKU_ID.AddItem CStr(lo.DataBodyRange(r, 1).Value)
        Next r
    End If
    If cbo_SKU_ID.ListCount > 0 Then cbo_SKU_ID.ListIndex = 0
    
    RechargerMois
    ActualiserPrevision
End Sub

Private Sub RechargerMois()
    cbo_Mois.Clear
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.sheets(SH_DEMAND_PLAN)
    Dim loDp As ListObject: Set loDp = wsDp.ListObjects("tbl_DEMAND_PLAN")
    If loDp Is Nothing Or loDp.DataBodyRange Is Nothing Then Exit Sub
    
    Dim dictMois As Object: Set dictMois = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To loDp.DataBodyRange.Rows.count
        Dim mVal As String: mVal = Format(loDp.DataBodyRange(r, 2).Value, "YYYY-MM")
        If Not dictMois.Exists(mVal) Then
            dictMois.Add mVal, 1
            cbo_Mois.AddItem mVal
        End If
    Next r
    If cbo_Mois.ListCount > 0 Then cbo_Mois.ListIndex = 0
End Sub

Private Sub cbo_SKU_ID_Change()
    ActualiserPrevision
End Sub

Private Sub cbo_Mois_Change()
    ActualiserPrevision
End Sub

Private Sub txt_Valeur_Ajustement_Change()
    CalculerProjection
End Sub

Private Sub cbo_Type_Ajustement_Change()
    CalculerProjection
End Sub

Private Sub ActualiserPrevision()
    If cbo_SKU_ID.ListIndex < 0 Or cbo_Mois.ListIndex < 0 Then Exit Sub
    Dim wsDp As Worksheet: Set wsDp = ThisWorkbook.sheets(SH_DEMAND_PLAN)
    Dim loDp As ListObject: Set loDp = wsDp.ListObjects("tbl_DEMAND_PLAN")
    If loDp Is Nothing Or loDp.DataBodyRange Is Nothing Then Exit Sub
    
    Dim r As Long
    For r = 1 To loDp.DataBodyRange.Rows.count
        Dim rSku As String: rSku = CStr(loDp.DataBodyRange(r, 1).Value)
        Dim rMois As String: rMois = Format(loDp.DataBodyRange(r, 2).Value, "YYYY-MM")
        If UCase(Trim(rSku)) = UCase(Trim(cbo_SKU_ID.Text)) And rMois = cbo_Mois.Text Then
            Dim statQte As Double: statQte = CDbl(loDp.DataBodyRange(r, 3).Value)
            Dim ajQte As Double: ajQte = CDbl(loDp.DataBodyRange(r, 4).Value)
            Dim finQte As Double: finQte = CDbl(loDp.DataBodyRange(r, 5).Value)
            
            lbl_ValPrevStat.Caption = Format(statQte, "#,##0") & " unit" & Chr(233) & "s"
            txt_Valeur_Ajustement.Text = CStr(ajQte)
            If loDp.ListColumns.count >= 7 Then txt_Motif_Ajustement.Text = CStr(loDp.DataBodyRange(r, 7).Value)
            CalculerProjection
            Exit Sub
        End If
    Next r
    lbl_ValPrevStat.Caption = "0 unit" & Chr(233) & "s"
    lbl_ValDemandeFinale.Caption = "0 unit" & Chr(233) & "s"
End Sub

Private Sub CalculerProjection()
    Dim statQte As Double: statQte = val(Replace(lbl_ValPrevStat.Caption, " ", ""))
    Dim valAj As Double: valAj = val(txt_Valeur_Ajustement.Text)
    Dim finQte As Double
    
    If InStr(1, cbo_Type_Ajustement.Text, "Pourcentage", vbTextCompare) > 0 Or InStr(1, cbo_Type_Ajustement.Text, "%") > 0 Then
        finQte = Round(statQte * (1# + (valAj / 100#)), 0)
    Else
        finQte = statQte + valAj
    End If
    If finQte < 0 Then finQte = 0
    lbl_ValDemandeFinale.Caption = Format(finQte, "#,##0") & " unit" & Chr(233) & "s"
End Sub

Private Sub btn_Fermer_Click()
    Unload Me
End Sub

Private Sub btn_ReinitialiserAjust_Click()
    txt_Valeur_Ajustement.Text = "0"
    txt_Motif_Ajustement.Text = "Remise " & Chr(224) & " z" & Chr(233) & "ro de l'ajustement commercial"
    btn_Enregistrer_Click
End Sub

Private Sub btn_Enregistrer_Click()
    If cbo_SKU_ID.ListIndex < 0 Or cbo_Mois.ListIndex < 0 Then
        MsgBox "Veuillez s" & Chr(233) & "lectionner un produit et un mois.", vbExclamation, TOOL_NAME: Exit Sub
    End If
    If Trim(txt_Motif_Ajustement.Text) = "" Then
        MsgBox "Le motif de l'ajustement est OBLIGATOIRE pour la tra" & Chr(231) & "abilit" & Chr(233) & " S&OP !" & vbCrLf & _
               "(Exemple: 'Promotion rentr" & Chr(233) & "e', 'Gain appel d'offres', 'P" & Chr(233) & "nurie concurrent')", vbExclamation, "Validation S&OP"
        txt_Motif_Ajustement.SetFocus: Exit Sub
    End If
    
    Dim ok As Boolean
    ok = AppliquerAjustementCommercial(cbo_SKU_ID.Text, cbo_Mois.Text, cbo_Type_Ajustement.Text, _
                                       val(txt_Valeur_Ajustement.Text), Trim(txt_Motif_Ajustement.Text))
    If ok Then
        MsgBox "Ajustement commercial enregistr" & Chr(233) & " avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        ActualiserPrevision
    End If
End Sub


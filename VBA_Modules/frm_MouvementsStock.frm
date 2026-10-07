VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_MouvementsStock 
   Caption         =   "Mouvements & Controle des Stocks (T3)"
   ClientHeight    =   6240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "frm_MouvementsStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_MouvementsStock"
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
    Me.Caption = "Mouvements & R" & Chr(233) & "gulation des Stocks (T3)"
    lbl_Title.Caption = "  SAISIE DES MOUVEMENTS & GESTION DES STOCKS"
    lbl_SKU.Caption = "Produit (SKU) :"
    lbl_Plant.Caption = "Usine :"
    lbl_StockActuel.Caption = "Stock Actuel :"
    lbl_ValStockActuel.Caption = "0 unit" & Chr(233) & "s"
    lbl_TypeMvt.Caption = "Type Mouvement :"
    lbl_QteMvt.Caption = "Quantit" & Chr(233) & " :"
    lbl_Motif.Caption = "Motif Mouvement :"
    lbl_Warning.Caption = "R" & Chr(232) & "gle stricte : aucun mouvement ne peut g" & Chr(233) & "n" & Chr(233) & "rer un stock n" & Chr(233) & "gatif."
    btn_ValiderMouvement.Caption = "Enregistrer"
    btn_Annuler.Caption = "R" & Chr(233) & "initialiser"
    btn_Fermer.Caption = "Fermer"
End Sub

Private Sub InitialiserFormulaire()
    cbo_Type_Mouvement.Clear
    cbo_Type_Mouvement.AddItem "Entr" & Chr(233) & "e R" & Chr(233) & "ception (+)"
    cbo_Type_Mouvement.AddItem "Sortie Livraison (-)"
    cbo_Type_Mouvement.AddItem "Ajustement Inventaire (+/-)"
    cbo_Type_Mouvement.AddItem "Perte DLC / Rebut (-)"
    cbo_Type_Mouvement.ListIndex = 0
    
    cbo_Plant_ID.Clear
    cbo_Plant_ID.AddItem "USN-01"
    cbo_Plant_ID.AddItem "USN-02"
    cbo_Plant_ID.AddItem "USN-03"
    cbo_Plant_ID.ListIndex = 0
    
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
    
    ActualiserStock
End Sub

Private Sub cbo_SKU_ID_Change()
    ActualiserStock
End Sub

Private Sub ActualiserStock()
    If cbo_SKU_ID.ListIndex >= 0 Then
        Dim stk As Double: stk = GetStockActuel(cbo_SKU_ID.Text)
        lbl_ValStockActuel.Caption = Format(stk, "#,##0") & " unit" & Chr(233) & "s"
        If stk <= 0 Then
            lbl_ValStockActuel.ForeColor = RGB(200, 0, 0)
        Else
            lbl_ValStockActuel.ForeColor = RGB(0, 100, 0)
        End If
    Else
        lbl_ValStockActuel.Caption = "0 unit" & Chr(233) & "s"
    End If
End Sub

Private Sub btn_Annuler_Click()
    txt_Quantite_Mouvement.Text = ""
    txt_Motif.Text = ""
    cbo_Type_Mouvement.ListIndex = 0
    ActualiserStock
End Sub

Private Sub btn_Fermer_Click()
    Unload Me
End Sub

Private Sub btn_ValiderMouvement_Click()
    If cbo_SKU_ID.ListIndex < 0 Then
        MsgBox "Veuillez s" & Chr(233) & "lectionner un produit (SKU).", vbExclamation, TOOL_NAME: Exit Sub
    End If
    If Not IsNumeric(txt_Quantite_Mouvement.Text) Or val(txt_Quantite_Mouvement.Text) <= 0 Then
        MsgBox "Veuillez saisir une quantit" & Chr(233) & " strictement positive.", vbExclamation, TOOL_NAME
        txt_Quantite_Mouvement.SetFocus: Exit Sub
    End If
    If Trim(txt_Motif.Text) = "" Then
        MsgBox "Le motif du mouvement est obligatoire.", vbExclamation, TOOL_NAME
        txt_Motif.SetFocus: Exit Sub
    End If
    
    Dim ok As Boolean
    ok = EnregistrerMouvementStock(cbo_SKU_ID.Text, cbo_Plant_ID.Text, cbo_Type_Mouvement.Text, _
                                  CDbl(txt_Quantite_Mouvement.Text), Trim(txt_Motif.Text))
    If ok Then
        MsgBox "Mouvement de stock enregistr" & Chr(233) & " avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        txt_Quantite_Mouvement.Text = ""
        txt_Motif.Text = ""
        ActualiserStock
    End If
End Sub


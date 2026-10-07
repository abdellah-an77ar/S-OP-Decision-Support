VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Produits 
   Caption         =   "Gestion des Produits & Master Data (T1)"
   ClientHeight    =   6240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "frm_Produits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Produits"
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
    Me.Caption = "Gestion des Produits & Master Data (T1)"
    lbl_Title.Caption = "  GESTION DU CATALOGUE PRODUITS (CRUD)"
    lbl_SKU.Caption = "Code SKU :"
    lbl_Nom.Caption = "D" & Chr(233) & "signation :"
    lbl_Famille.Caption = "Famille :"
    lbl_Plant.Caption = "Usine :"
    lbl_Poids.Caption = "Poids Unit (kg) :"
    lbl_DLC.Caption = "DLC (jours) :"
    lbl_Prix.Caption = "Prix Vente :"
    lbl_Cout.Caption = "Co" & Chr(251) & "t Std :"
    chk_Statut_Actif.Caption = "Produit Actif (d" & Chr(233) & "cocher pour d" & Chr(233) & "sactiver)"
    lbl_Info.Caption = "Pr" & Chr(234) & "t pour saisie ou recherche."
    btn_Ajouter.Caption = "Ajouter"
    btn_Rechercher.Caption = "Chercher"
    btn_Modifier.Caption = "Modifier"
    btn_Desactiver.Caption = "D" & Chr(233) & "sactiver"
    btn_Vider.Caption = "Nouveau"
    btn_Fermer.Caption = "Fermer"
End Sub

Private Sub InitialiserFormulaire()
    cbo_Famille.Clear
    cbo_Famille.AddItem "Biscuits"
    cbo_Famille.AddItem "Crackers"
    cbo_Famille.AddItem "Gaufrettes"
    cbo_Famille.AddItem "Couscous"
    cbo_Famille.AddItem "P" & Chr(226) & "tes"
    cbo_Famille.AddItem "Farine"
    cbo_Famille.AddItem "Huile"
    cbo_Famille.ListIndex = 0
    
    cbo_Plant_ID.Clear
    cbo_Plant_ID.AddItem "USN-01"
    cbo_Plant_ID.AddItem "USN-02"
    cbo_Plant_ID.AddItem "USN-03"
    cbo_Plant_ID.ListIndex = 0
    
    RechargerListeProduits
    ViderChamps
End Sub

Private Sub RechargerListeProduits()
    On Error Resume Next
    cbo_SelectSKU.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.sheets(SH_PRODUCT)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = ws.ListObjects("tbl_PRODUCT")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        cbo_SelectSKU.AddItem CStr(lo.DataBodyRange(r, 1).Value)
    Next r
End Sub

Private Sub ViderChamps()
    txt_SKU_ID.Text = ""
    txt_Nom.Text = ""
    cbo_Famille.ListIndex = 0
    cbo_Plant_ID.ListIndex = 0
    txt_Poids_unit.Text = "1.00"
    txt_DLC_jours.Text = "180"
    txt_Prix_vente.Text = "0"
    txt_Cout_standard.Text = "0"
    chk_Statut_Actif.Value = True
    lbl_Info.Caption = "Pr" & Chr(234) & "t pour saisie ou recherche."
    lbl_Info.ForeColor = RGB(100, 100, 100)
    txt_SKU_ID.Enabled = True
    txt_SKU_ID.SetFocus
End Sub

Private Sub btn_Vider_Click()
    ViderChamps
End Sub

Private Sub btn_Fermer_Click()
    Unload Me
End Sub

Private Sub cbo_SelectSKU_Change()
    If cbo_SelectSKU.ListIndex >= 0 Then
        txt_SKU_ID.Text = cbo_SelectSKU.Text
        RechercherSKU cbo_SelectSKU.Text
    End If
End Sub

Private Sub btn_Rechercher_Click()
    If Trim(txt_SKU_ID.Text) = "" Then
        MsgBox "Veuillez saisir un code SKU " & Chr(224) & " rechercher.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    RechercherSKU Trim(txt_SKU_ID.Text)
End Sub

Private Sub RechercherSKU(ByVal sku As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.sheets(SH_PRODUCT)
    Dim lo As ListObject: Set lo = ws.ListObjects("tbl_PRODUCT")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    
    Dim r As Long, found As Boolean: found = False
    For r = 1 To lo.DataBodyRange.Rows.count
        If UCase(Trim(CStr(lo.DataBodyRange(r, 1).Value))) = UCase(Trim(sku)) Then
            txt_SKU_ID.Text = CStr(lo.DataBodyRange(r, 1).Value)
            txt_Nom.Text = CStr(lo.DataBodyRange(r, 2).Value)
            cbo_Famille.Text = CStr(lo.DataBodyRange(r, 3).Value)
            cbo_Plant_ID.Text = CStr(lo.DataBodyRange(r, 4).Value)
            txt_Poids_unit.Text = Format(lo.DataBodyRange(r, 5).Value, "0.00")
            txt_DLC_jours.Text = CStr(lo.DataBodyRange(r, 6).Value)
            If lo.ListColumns.count >= 8 Then txt_Prix_vente.Text = CStr(lo.DataBodyRange(r, 8).Value) Else txt_Prix_vente.Text = "0"
            If lo.ListColumns.count >= 9 Then txt_Cout_standard.Text = CStr(lo.DataBodyRange(r, 9).Value) Else txt_Cout_standard.Text = "0"
            If lo.ListColumns.count >= 10 Then
                chk_Statut_Actif.Value = (UCase(Trim(CStr(lo.DataBodyRange(r, 10).Value))) = "OUI")
            Else
                chk_Statut_Actif.Value = True
            End If
            
            lbl_Info.Caption = "Produit charg" & Chr(233) & " : " & sku
            lbl_Info.ForeColor = RGB(0, 120, 0)
            txt_SKU_ID.Enabled = False
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        MsgBox "Le produit SKU '" & sku & "' n'a pas " & Chr(233) & "t" & Chr(233) & " trouv" & Chr(233) & " dans le catalogue.", vbInformation, TOOL_NAME
        lbl_Info.Caption = "Produit non trouv" & Chr(233) & "."
        lbl_Info.ForeColor = RGB(180, 0, 0)
    End If
End Sub

Private Sub btn_Ajouter_Click()
    If Not ValiderChamps() Then Exit Sub
    Dim ok As Boolean
    ok = SauvegarderProduit(Trim(txt_SKU_ID.Text), Trim(txt_Nom.Text), cbo_Famille.Text, cbo_Plant_ID.Text, _
                            CDbl(txt_Poids_unit.Text), CLng(txt_DLC_jours.Text), _
                            CDbl(txt_Prix_vente.Text), CDbl(txt_Cout_standard.Text), chk_Statut_Actif.Value, True)
    If ok Then
        MsgBox "Produit '" & Trim(txt_SKU_ID.Text) & "' ajout" & Chr(233) & " avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        RechargerListeProduits
        ViderChamps
    End If
End Sub

Private Sub btn_Modifier_Click()
    If Not ValiderChamps() Then Exit Sub
    Dim ok As Boolean
    ok = SauvegarderProduit(Trim(txt_SKU_ID.Text), Trim(txt_Nom.Text), cbo_Famille.Text, cbo_Plant_ID.Text, _
                            CDbl(txt_Poids_unit.Text), CLng(txt_DLC_jours.Text), _
                            CDbl(txt_Prix_vente.Text), CDbl(txt_Cout_standard.Text), chk_Statut_Actif.Value, False)
    If ok Then
        MsgBox "Produit '" & Trim(txt_SKU_ID.Text) & "' modifi" & Chr(233) & " avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        RechargerListeProduits
        ViderChamps
    End If
End Sub

Private Sub btn_Desactiver_Click()
    If Trim(txt_SKU_ID.Text) = "" Then
        MsgBox "Veuillez s" & Chr(233) & "lectionner un produit " & Chr(224) & " d" & Chr(233) & "sactiver.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    If MsgBox("Voulez-vous d" & Chr(233) & "sactiver le produit '" & txt_SKU_ID.Text & "' ?" & vbCrLf & _
              "(L'historique sera conserv" & Chr(233) & " mais le produit ne sera plus commandable)", _
              vbQuestion + vbYesNo, TOOL_NAME) = vbYes Then
        Dim ok As Boolean: ok = DesactiverProduit(Trim(txt_SKU_ID.Text))
        If ok Then
            MsgBox "Produit d" & Chr(233) & "sactiv" & Chr(233) & " avec succ" & Chr(232) & "s.", vbInformation, TOOL_NAME
            RechargerListeProduits
            ViderChamps
        End If
    End If
End Sub

Private Function ValiderChamps() As Boolean
    ValiderChamps = False
    If Trim(txt_SKU_ID.Text) = "" Then
        MsgBox "Le code SKU est obligatoire.", vbExclamation, TOOL_NAME: txt_SKU_ID.SetFocus: Exit Function
    End If
    If Trim(txt_Nom.Text) = "" Then
        MsgBox "La d" & Chr(233) & "signation du produit est obligatoire.", vbExclamation, TOOL_NAME: txt_Nom.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_Poids_unit.Text) Or val(txt_Poids_unit.Text) <= 0 Then
        MsgBox "Poids unitaire invalide (> 0).", vbExclamation, TOOL_NAME: txt_Poids_unit.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_DLC_jours.Text) Or val(txt_DLC_jours.Text) <= 0 Then
        MsgBox "DLC invalide (> 0 jours).", vbExclamation, TOOL_NAME: txt_DLC_jours.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_Prix_vente.Text) Or val(txt_Prix_vente.Text) < 0 Then
        MsgBox "Prix de vente invalide.", vbExclamation, TOOL_NAME: txt_Prix_vente.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_Cout_standard.Text) Or val(txt_Cout_standard.Text) < 0 Then
        MsgBox "Co" & Chr(251) & "t standard invalide.", vbExclamation, TOOL_NAME: txt_Cout_standard.SetFocus: Exit Function
    End If
    ValiderChamps = True
End Function


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Commandes 
   Caption         =   "Suivi des Commandes Clients (T2)"
   ClientHeight    =   6240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "frm_Commandes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Commandes"
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
    Me.Caption = "Suivi des Commandes Clients (T2)"
    lbl_Title.Caption = "  SUIVI & SAISIE DES COMMANDES CLIENTS"
    lbl_CmdID.Caption = "N" & Chr(176) & " Commande :"
    btn_Rechercher.Caption = "Chercher"
    lbl_Client.Caption = "Client / Canal :"
    lbl_SKU.Caption = "Produit (SKU) :"
    lbl_Qte.Caption = "Quantit" & Chr(233) & " :"
    lbl_DateCmd.Caption = "Date Cmd :"
    lbl_DatePrev.Caption = "Livraison Pr" & Chr(233) & "vue :"
    lbl_DateReel.Caption = "Livraison R" & Chr(233) & "elle :"
    lbl_Statut.Caption = "Statut :"
    lbl_Comment.Caption = "Commentaire :"
    lbl_Info.Caption = "Pr" & Chr(234) & "t pour saisie ou recherche."
    btn_Creer.Caption = "Cr" & Chr(233) & "er Cmd"
    btn_MettreAJour.Caption = "Mettre " & Chr(224) & " Jour"
    btn_Annuler.Caption = "Annuler Cmd"
    btn_Vider.Caption = "Nouveau"
    btn_Fermer.Caption = "Fermer"
End Sub

Private Sub InitialiserFormulaire()
    cbo_Statut.Clear
    cbo_Statut.AddItem "En cours"
    cbo_Statut.AddItem "Livr" & Chr(233) & "e"
    cbo_Statut.AddItem "Livr" & Chr(233) & "e en retard"
    cbo_Statut.AddItem "En retard"
    cbo_Statut.AddItem "Annul" & Chr(233) & "e"
    cbo_Statut.ListIndex = 0
    
    cbo_SKU_ID.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.sheets(SH_PRODUCT)
    Dim lo As ListObject: Set lo = ws.ListObjects("tbl_PRODUCT")
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.count
            cbo_SKU_ID.AddItem CStr(lo.DataBodyRange(r, 1).Value)
        Next r
    End If
    
    RechargerCommandes
    ViderChamps
End Sub

Private Sub RechargerCommandes()
    On Error Resume Next
    cbo_SelectCmd.Clear
    Dim wsCmd As Worksheet: Set wsCmd = ThisWorkbook.sheets("COMMANDES")
    If wsCmd Is Nothing Then Exit Sub
    Dim loCmd As ListObject: Set loCmd = wsCmd.ListObjects("tbl_COMMANDES")
    If loCmd Is Nothing Or loCmd.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long
    For r = 1 To loCmd.DataBodyRange.Rows.count
        cbo_SelectCmd.AddItem CStr(loCmd.DataBodyRange(r, 1).Value)
    Next r
End Sub

Private Sub ViderChamps()
    txt_Commande_ID.Text = "CMD-" & Format(Now, "YYYYMMDD-HHNN")
    txt_Client.Text = ""
    If cbo_SKU_ID.ListCount > 0 Then cbo_SKU_ID.ListIndex = 0
    txt_Quantite.Text = "100"
    txt_Date_Commande.Text = Format(Date, "YYYY-MM-DD")
    txt_Date_Livraison_Prevue.Text = Format(Date + 5, "YYYY-MM-DD")
    txt_Date_Livraison_Reelle.Text = ""
    cbo_Statut.ListIndex = 0
    txt_Commentaire.Text = ""
    lbl_Info.Caption = "Pr" & Chr(234) & "t pour saisie ou recherche."
    lbl_Info.ForeColor = RGB(100, 100, 100)
    txt_Commande_ID.Enabled = True
    txt_Client.SetFocus
End Sub

Private Sub btn_Vider_Click()
    ViderChamps
End Sub

Private Sub btn_Fermer_Click()
    Unload Me
End Sub

Private Sub cbo_SelectCmd_Change()
    If cbo_SelectCmd.ListIndex >= 0 Then
        RechercherCommande cbo_SelectCmd.Text
    End If
End Sub

Private Sub btn_Rechercher_Click()
    If Trim(txt_Commande_ID.Text) = "" Then
        MsgBox "Veuillez renseigner un N" & Chr(176) & " de commande.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    RechercherCommande Trim(txt_Commande_ID.Text)
End Sub

Private Sub RechercherCommande(ByVal cmdId As String)
    Dim wsCmd As Worksheet: Set wsCmd = ThisWorkbook.sheets("COMMANDES")
    Dim loCmd As ListObject: Set loCmd = wsCmd.ListObjects("tbl_COMMANDES")
    If loCmd Is Nothing Or loCmd.DataBodyRange Is Nothing Then Exit Sub
    
    Dim r As Long, found As Boolean: found = False
    For r = 1 To loCmd.DataBodyRange.Rows.count
        If UCase(Trim(CStr(loCmd.DataBodyRange(r, 1).Value))) = UCase(Trim(cmdId)) Then
            txt_Commande_ID.Text = CStr(loCmd.DataBodyRange(r, 1).Value)
            txt_Client.Text = CStr(loCmd.DataBodyRange(r, 2).Value)
            cbo_SKU_ID.Text = CStr(loCmd.DataBodyRange(r, 3).Value)
            txt_Quantite.Text = CStr(loCmd.DataBodyRange(r, 4).Value)
            txt_Date_Commande.Text = Format(loCmd.DataBodyRange(r, 5).Value, "YYYY-MM-DD")
            txt_Date_Livraison_Prevue.Text = Format(loCmd.DataBodyRange(r, 6).Value, "YYYY-MM-DD")
            If Not IsEmpty(loCmd.DataBodyRange(r, 7).Value) Then
                txt_Date_Livraison_Reelle.Text = Format(loCmd.DataBodyRange(r, 7).Value, "YYYY-MM-DD")
            Else
                txt_Date_Livraison_Reelle.Text = ""
            End If
            cbo_Statut.Text = CStr(loCmd.DataBodyRange(r, 8).Value)
            If loCmd.ListColumns.count >= 9 Then txt_Commentaire.Text = CStr(loCmd.DataBodyRange(r, 9).Value)
            
            lbl_Info.Caption = "Commande charg" & Chr(233) & "e : " & cmdId
            lbl_Info.ForeColor = RGB(0, 120, 0)
            txt_Commande_ID.Enabled = False
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        MsgBox "Commande '" & cmdId & "' introuvable.", vbInformation, TOOL_NAME
        lbl_Info.Caption = "Commande introuvable."
        lbl_Info.ForeColor = RGB(180, 0, 0)
    End If
End Sub

Private Sub btn_Creer_Click()
    If Not ValiderChamps() Then Exit Sub
    Dim ok As Boolean
    ok = SauvegarderCommande(Trim(txt_Commande_ID.Text), Trim(txt_Client.Text), cbo_SKU_ID.Text, _
                             CDbl(txt_Quantite.Text), Trim(txt_Date_Commande.Text), _
                             Trim(txt_Date_Livraison_Prevue.Text), Trim(txt_Date_Livraison_Reelle.Text), _
                             cbo_Statut.Text, Trim(txt_Commentaire.Text), True)
    If ok Then
        MsgBox "Commande '" & Trim(txt_Commande_ID.Text) & "' enregistr" & Chr(233) & "e avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        RechargerCommandes
        ViderChamps
    End If
End Sub

Private Sub btn_MettreAJour_Click()
    If Not ValiderChamps() Then Exit Sub
    Dim ok As Boolean
    ok = SauvegarderCommande(Trim(txt_Commande_ID.Text), Trim(txt_Client.Text), cbo_SKU_ID.Text, _
                             CDbl(txt_Quantite.Text), Trim(txt_Date_Commande.Text), _
                             Trim(txt_Date_Livraison_Prevue.Text), Trim(txt_Date_Livraison_Reelle.Text), _
                             cbo_Statut.Text, Trim(txt_Commentaire.Text), False)
    If ok Then
        MsgBox "Commande '" & Trim(txt_Commande_ID.Text) & "' mise " & Chr(224) & " jour avec succ" & Chr(232) & "s !", vbInformation, TOOL_NAME
        RechargerCommandes
        ViderChamps
    End If
End Sub

Private Sub btn_Annuler_Click()
    If Trim(txt_Commande_ID.Text) = "" Then
        MsgBox "Veuillez s" & Chr(233) & "lectionner une commande " & Chr(224) & " annuler.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    If MsgBox("Voulez-vous annuler la commande '" & txt_Commande_ID.Text & "' ?", vbQuestion + vbYesNo, TOOL_NAME) = vbYes Then
        Dim ok As Boolean: ok = AnnulerCommande(Trim(txt_Commande_ID.Text))
        If ok Then
            MsgBox "Commande annul" & Chr(233) & "e avec succ" & Chr(232) & "s.", vbInformation, TOOL_NAME
            RechargerCommandes
            ViderChamps
        End If
    End If
End Sub

Private Function ValiderChamps() As Boolean
    ValiderChamps = False
    If Trim(txt_Commande_ID.Text) = "" Then
        MsgBox "N" & Chr(176) & " Commande obligatoire.", vbExclamation, TOOL_NAME: txt_Commande_ID.SetFocus: Exit Function
    End If
    If Trim(txt_Client.Text) = "" Then
        MsgBox "Nom du Client obligatoire.", vbExclamation, TOOL_NAME: txt_Client.SetFocus: Exit Function
    End If
    If cbo_SKU_ID.ListIndex < 0 Then
        MsgBox "Veuillez s" & Chr(233) & "lectionner un produit (SKU).", vbExclamation, TOOL_NAME: cbo_SKU_ID.SetFocus: Exit Function
    End If
    If Not IsNumeric(txt_Quantite.Text) Or val(txt_Quantite.Text) <= 0 Then
        MsgBox "Quantit" & Chr(233) & " invalide (doit " & Chr(234) & "tre > 0).", vbExclamation, TOOL_NAME: txt_Quantite.SetFocus: Exit Function
    End If
    If Not IsDate(txt_Date_Commande.Text) Then
        MsgBox "Date de commande invalide (format attendu YYYY-MM-DD).", vbExclamation, TOOL_NAME: txt_Date_Commande.SetFocus: Exit Function
    End If
    If Not IsDate(txt_Date_Livraison_Prevue.Text) Then
        MsgBox "Date de livraison pr" & Chr(233) & "vue invalide (format attendu YYYY-MM-DD).", vbExclamation, TOOL_NAME: txt_Date_Livraison_Prevue.SetFocus: Exit Function
    End If
    If CDate(txt_Date_Livraison_Prevue.Text) < CDate(txt_Date_Commande.Text) Then
        MsgBox "La date de livraison ne peut pas " & Chr(234) & "tre ant" & Chr(233) & "rieure " & Chr(224) & " la commande !", vbExclamation, TOOL_NAME: Exit Function
    End If
    ValiderChamps = True
End Function


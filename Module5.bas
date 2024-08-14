Attribute VB_Name = "Module5"
Sub VerifierEtRemplirDecision()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim decisionCol As Range

    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    ' Trouver la dernière ligne avec des données dans la colonne AN (ou une autre des colonnes concernées)
    lastRow = ws.Cells(ws.Rows.Count, "AN").End(xlUp).Row
    
    ' Définir la nouvelle colonne pour la décision
    Set decisionCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1)
    
    ' Nommer la nouvelle colonne "decision"
    decisionCol.Value = "decision"
    
    ' Boucler à travers chaque ligne pour vérifier les conditions
    For i = 5 To lastRow ' Commence à la ligne 5 pour ignorer les en-têtes
        If ((ws.Cells(i, "AN").Value = "21" Or ws.Cells(i, "AO").Value = "21") And _
             (ws.Cells(i, "AP").Value = "21" And ws.Cells(i, "AQ").Value = "21")) Then
            ws.Cells(i, decisionCol.Column).Value = "21P"
        Else
            ws.Cells(i, decisionCol.Column).Value = "x"
        End If
    Next i
    
    MsgBox "Vérification et remplissage de la colonne 'decision' terminés!"

End Sub



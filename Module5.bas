Attribute VB_Name = "Module5"
Sub VerifierEtRemplirDecision()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim decisionCol As Range

    ' D�finir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    ' Trouver la derni�re ligne avec des donn�es dans la colonne AN (ou une autre des colonnes concern�es)
    lastRow = ws.Cells(ws.Rows.Count, "AN").End(xlUp).Row
    
    ' D�finir la nouvelle colonne pour la d�cision
    Set decisionCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1)
    
    ' Nommer la nouvelle colonne "decision"
    decisionCol.Value = "decision"
    
    ' Boucler � travers chaque ligne pour v�rifier les conditions
    For i = 5 To lastRow ' Commence � la ligne 5 pour ignorer les en-t�tes
        If ((ws.Cells(i, "AN").Value = "21" Or ws.Cells(i, "AO").Value = "21") And _
             (ws.Cells(i, "AP").Value = "21" And ws.Cells(i, "AQ").Value = "21")) Then
            ws.Cells(i, decisionCol.Column).Value = "21P"
        Else
            ws.Cells(i, decisionCol.Column).Value = "x"
        End If
    Next i
    
    MsgBox "V�rification et remplissage de la colonne 'decision' termin�s!"

End Sub



Attribute VB_Name = "Module7"
Sub ColorerLignes()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim decisionCol As Range
    Dim apCol As Range
    Dim aqCol As Range

    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    ' Trouver la dernière ligne avec des données dans la colonne AR
    lastRow = ws.Cells(ws.Rows.Count, "AR").End(xlUp).Row

    ' Boucler à travers chaque ligne
    For i = 9 To lastRow ' Commence à la ligne 9 si les premières lignes sont des en-têtes
        Set decisionCol = ws.Cells(i, "AR")
        Set apCol = ws.Cells(i, "AP")
        Set aqCol = ws.Cells(i, "AQ")
        
        ' Vérifier si AP et AQ contiennent "21"
        If apCol.Value = "21" And aqCol.Value = "21" And decisionCol.Value <> "21P" Then
            ' Colorer toute la ligne en bleu
            ws.Rows(i).Interior.Color = RGB(173, 216, 230)
        ' Sinon, vérifier si la colonne "decision" contient "21P"
        ElseIf decisionCol.Value = "21P" Then
            ' Colorer toute la ligne en vert
            ws.Rows(i).Interior.Color = RGB(198, 224, 180)
        End If
    Next i

    MsgBox "Coloration des lignes terminée!"
    
End Sub


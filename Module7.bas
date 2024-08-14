Attribute VB_Name = "Module7"
Sub ColorerLignes()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim decisionCol As Range
    Dim apCol As Range
    Dim aqCol As Range

    ' D�finir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    ' Trouver la derni�re ligne avec des donn�es dans la colonne AR
    lastRow = ws.Cells(ws.Rows.Count, "AR").End(xlUp).Row

    ' Boucler � travers chaque ligne
    For i = 9 To lastRow ' Commence � la ligne 9 si les premi�res lignes sont des en-t�tes
        Set decisionCol = ws.Cells(i, "AR")
        Set apCol = ws.Cells(i, "AP")
        Set aqCol = ws.Cells(i, "AQ")
        
        ' V�rifier si AP et AQ contiennent "21"
        If apCol.Value = "21" And aqCol.Value = "21" And decisionCol.Value <> "21P" Then
            ' Colorer toute la ligne en bleu
            ws.Rows(i).Interior.Color = RGB(173, 216, 230)
        ' Sinon, v�rifier si la colonne "decision" contient "21P"
        ElseIf decisionCol.Value = "21P" Then
            ' Colorer toute la ligne en vert
            ws.Rows(i).Interior.Color = RGB(198, 224, 180)
        End If
    Next i

    MsgBox "Coloration des lignes termin�e!"
    
End Sub


Attribute VB_Name = "Module6"
Sub CreerColonnePrediction4()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nouvelleColonne As Range
    
    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille
    
    ' Trouver la dernière ligne avec des données dans la colonne L
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
   ' Définir la nouvelle colonne pour la prédiction
    Set nouvelleColonne = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1)
    
    ' Nommer la nouvelle colonne
    nouvelleColonne.Value = "HTUINFO"
    
    ' Boucler à travers chaque ligne et ajouter les valeurs de la prédiction
    For i = 9 To lastRow ' Commence à la ligne 2 si la première ligne est un en-tête
        If ws.Cells(i, "U").Value < ws.Cells(i, "T").Value Then
            nouvelleColonne.Offset(i - 1, 0).Value = "21"
        Else
            nouvelleColonne.Offset(i - 1, 0).Value = "x"
        End If
    Next i
    
    MsgBox "Colonne 'A1419' créée et remplie!"
    
End Sub





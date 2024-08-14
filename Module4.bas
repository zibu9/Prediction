Attribute VB_Name = "Module4"
Sub CreerColonnePrediction3()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nouvelleColonne As Range
    
    ' D�finir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne L
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
   ' D�finir la nouvelle colonne pour la pr�diction
    Set nouvelleColonne = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1)
    
    ' Nommer la nouvelle colonne
    nouvelleColonne.Value = "HTU146"
    
    ' Boucler � travers chaque ligne et ajouter les valeurs de la pr�diction
    For i = 9 To lastRow ' Commence � la ligne 2 si la premi�re ligne est un en-t�te
        If ws.Cells(i, "U").Value < 1.46 Then
            nouvelleColonne.Offset(i - 1, 0).Value = "21"
        Else
            nouvelleColonne.Offset(i - 1, 0).Value = "x"
        End If
    Next i
    
    MsgBox "Colonne 'A1419' cr��e et remplie!"
    
End Sub




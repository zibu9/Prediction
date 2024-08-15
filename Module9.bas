Attribute VB_Name = "Module9"
Sub trierVingt()

    Dim ws As Worksheet, wsVingt As Worksheet
    Dim lastRow As Long, vingtRow As Long
    Dim i As Long
    Dim col As Range

    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("VingtPropable").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Créer la nouvelle feuille
    Set wsVingt = ThisWorkbook.Sheets.Add(After:=ws)
    wsVingt.Name = "VingtPropable"

    ' Copier les en-têtes de la feuille principale
    ws.Range("A1:AR8").Copy Destination:=wsVingt.Range("A1")
    
    ' Copier la taille des colonnes de la feuille principale vers la nouvelle feuille
    For Each col In ws.Range("A1:AR1").Columns
        wsVingt.Columns(col.Column).ColumnWidth = col.ColumnWidth
    Next col

    ' Initialiser les lignes pour la nouvelle feuille
    vingtRow = 9

    ' Trouver la dernière ligne avec des données dans la colonne AR
    lastRow = ws.Cells(ws.Rows.Count, "AR").End(xlUp).Row

    ' Boucler à travers chaque ligne
    For i = 9 To lastRow ' Commence à la ligne 9

        ' Vérifier si les colonnes J, L, U et T remplissent les conditions
        If ((ws.Cells(i, "J").Value < 1.27 Or ws.Cells(i, "L").Value < 1.27) And _
             (ws.Cells(i, "U").Value > 0 And ws.Cells(i, "T").Value < 2.2)) Then
            ' Copier la ligne dans la feuille "VingtPropable"
            ws.Range("A" & i & ":AR" & i).Copy Destination:=wsVingt.Range("A" & vingtRow)
            ' Colorer la ligne en blanc dans la feuille "VingtPropable"
            wsVingt.Rows(vingtRow).Interior.Color = RGB(255, 255, 255)
            vingtRow = vingtRow + 1
        End If
    Next i

    MsgBox "Création de la feuille 20 terminée!"
    
End Sub



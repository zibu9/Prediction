Attribute VB_Name = "Module9"
Sub trierVingt()

    Dim ws As Worksheet, wsVingt As Worksheet
    Dim lastRow As Long, vingtRow As Long
    Dim i As Long
    Dim col As Range

    ' D�finir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("VingtPropable").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Cr�er la nouvelle feuille
    Set wsVingt = ThisWorkbook.Sheets.Add(After:=ws)
    wsVingt.Name = "VingtPropable"

    ' Copier les en-t�tes de la feuille principale
    ws.Range("A1:AR8").Copy Destination:=wsVingt.Range("A1")
    
    ' Copier la taille des colonnes de la feuille principale vers la nouvelle feuille
    For Each col In ws.Range("A1:AR1").Columns
        wsVingt.Columns(col.Column).ColumnWidth = col.ColumnWidth
    Next col

    ' Initialiser les lignes pour la nouvelle feuille
    vingtRow = 9

    ' Trouver la derni�re ligne avec des donn�es dans la colonne AR
    lastRow = ws.Cells(ws.Rows.Count, "AR").End(xlUp).Row

    ' Boucler � travers chaque ligne
    For i = 9 To lastRow ' Commence � la ligne 9

        ' V�rifier si les colonnes J, L, U et T remplissent les conditions
        If ((ws.Cells(i, "J").Value < 1.27 Or ws.Cells(i, "L").Value < 1.27) And _
             (ws.Cells(i, "U").Value > 0 And ws.Cells(i, "T").Value < 2.2)) Then
            ' Copier la ligne dans la feuille "VingtPropable"
            ws.Range("A" & i & ":AR" & i).Copy Destination:=wsVingt.Range("A" & vingtRow)
            ' Colorer la ligne en blanc dans la feuille "VingtPropable"
            wsVingt.Rows(vingtRow).Interior.Color = RGB(255, 255, 255)
            vingtRow = vingtRow + 1
        End If
    Next i

    MsgBox "Cr�ation de la feuille 20 termin�e!"
    
End Sub



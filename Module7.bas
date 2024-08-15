Attribute VB_Name = "Module7"
Sub ColorerLignes()

    Dim ws As Worksheet, wsVert As Worksheet, wsBleu As Worksheet
    Dim lastRow As Long, vertRow As Long, bleuRow As Long
    Dim i As Long
    Dim decisionCol As Range, apCol As Range, aqCol As Range

    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille

    ' Supprimer les feuilles "Vert" et "Bleu" si elles existent déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Vert").Delete
    ThisWorkbook.Sheets("Bleu").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Créer les nouvelles feuilles "Vert" et "Bleu"
    Set wsVert = ThisWorkbook.Sheets.Add(After:=ws)
    wsVert.Name = "Vert"
    Set wsBleu = ThisWorkbook.Sheets.Add(After:=ws)
    wsBleu.Name = "Bleu"

    ' Copier les en-têtes de la feuille principale vers les feuilles "Vert" et "Bleu"
    ws.Rows(1).Copy Destination:=wsVert.Rows(1)
    ws.Rows(1).Copy Destination:=wsBleu.Rows(1)
    
    ' Initialiser les lignes pour les nouvelles feuilles
    vertRow = 2
    bleuRow = 2

    ' Trouver la dernière ligne avec des données dans la colonne AR
    lastRow = ws.Cells(ws.Rows.Count, "AR").End(xlUp).Row

    ' Boucler à travers chaque ligne
    For i = 9 To lastRow ' Commence à la ligne 9 si les premières lignes sont des en-têtes
        Set decisionCol = ws.Cells(i, "AR")
        Set apCol = ws.Cells(i, "AP")
        Set aqCol = ws.Cells(i, "AQ")
        
        ' Vérifier si la colonne "decision" contient "21P"
        If decisionCol.Value = "21P" Then
            ' Colorer toute la ligne en vert
            ws.Rows(i).Interior.Color = RGB(198, 224, 180)
            ' Copier la ligne dans la feuille "Vert"
            ws.Rows(i).Copy Destination:=wsVert.Rows(vertRow)
            vertRow = vertRow + 1
        End If

        ' Vérifier si les colonnes AP et AQ contiennent "21"
        If apCol.Value = "21" And aqCol.Value = "21" Then
            ' Colorer toute la ligne en bleu
            ws.Rows(i).Interior.Color = RGB(173, 216, 230)
            ' Copier la ligne dans la feuille "Bleu"
            ws.Rows(i).Copy Destination:=wsBleu.Rows(bleuRow)
            bleuRow = bleuRow + 1
        End If
    Next i

    MsgBox "Coloration et création des nouvelles feuilles terminées!"
    
End Sub



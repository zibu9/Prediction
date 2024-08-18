Attribute VB_Name = "Module10"
Sub TirageAleatoire(sheetName As String, lineCount As Integer)

    Dim wsSource As Worksheet
    Dim wsTirage As Worksheet
    Dim tirageRow As Long
    Dim i As Long
    Dim randIndex As Long
    Dim usedIndexes As Collection
    Dim lastRow As Long
    Dim rowIndexes() As Long
    Dim col As Range

    ' Définir la feuille source
    Set wsSource = ThisWorkbook.Sheets(sheetName)
    
    ' Supprimer la feuille de tirage si elle existe déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Tirage_" & lineCount & "_" & sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Créer la nouvelle feuille de tirage
    Set wsTirage = ThisWorkbook.Sheets.Add(After:=wsSource)
    wsTirage.Name = "Tirage_" & lineCount & "_" & sheetName

    ' Copier les en-têtes de la feuille source
    wsSource.Range("A1:AR8").Copy Destination:=wsTirage.Range("A1")
    
    ' Copier la taille des colonnes de la feuille source vers la nouvelle feuille
    For Each col In wsSource.Range("A1:AR1").Columns
        wsTirage.Columns(col.Column).ColumnWidth = col.ColumnWidth
    Next col

    ' Trouver la dernière ligne avec des données dans la feuille source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Initialiser le tableau pour stocker les index des lignes
    ReDim rowIndexes(1 To lastRow - 8)

    ' Remplir le tableau avec les index des lignes à partir de la ligne 9
    For i = 9 To lastRow
        rowIndexes(i - 8) = i
    Next i

    ' Initialiser la collection pour suivre les index déjà utilisés
    Set usedIndexes = New Collection

    ' Initialiser la ligne de départ pour la feuille de tirage
    tirageRow = 9

    ' Tirer les lignes aléatoirement
    For i = 1 To lineCount
        Do
            randIndex = rowIndexes(Application.WorksheetFunction.RandBetween(1, UBound(rowIndexes)))
        Loop While ItemExists(usedIndexes, randIndex)
        
        ' Copier la ligne sélectionnée dans la feuille de tirage
        wsSource.Rows(randIndex).Copy Destination:=wsTirage.Rows(tirageRow)
        tirageRow = tirageRow + 1
        
        ' Marquer cet index comme utilisé
        usedIndexes.Add randIndex
    Next i

    MsgBox "Tirage aléatoire de " & lineCount & " lignes terminé!"

End Sub

' Fonction pour vérifier si un élément existe dans une collection
Function ItemExists(col As Collection, item As Variant) As Boolean
    On Error Resume Next
    ItemExists = Not IsError(col(item))
    On Error GoTo 0
End Function


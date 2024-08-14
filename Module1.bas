Attribute VB_Name = "Module1"
Sub ConvertirChaineEnNombre()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim colonnes As Variant
    Dim valeurNettoyee As String

    ' Définir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille
    
    ' Colonnes à traiter
    colonnes = Array("J", "K", "L", "T", "U")
    
    ' Trouver la dernière ligne avec des données dans la colonne J (ou la plus longue des colonnes concernées)
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Boucler à travers chaque ligne
    For i = 9 To lastRow ' Commence à la ligne 9 comme indiqué
    
        ' Boucler à travers chaque colonne à traiter
        For j = LBound(colonnes) To UBound(colonnes)
            With ws.Cells(i, colonnes(j))
                ' Nettoyer la valeur : enlever les espaces et remplacer le point par une virgule
                valeurNettoyee = Trim(.Value) ' Enlever les espaces
                valeurNettoyee = Replace(valeurNettoyee, ".", ",") ' Remplacer le point par une virgule
                
                ' Vérifier si la cellule contient une chaîne de caractère numérique et convertir
                If IsNumeric(valeurNettoyee) Then
                    .Value = CDbl(valeurNettoyee) ' Convertir en double
                End If
            End With
        Next j
        
    Next i
    
    MsgBox "Conversion des chaînes de caractères en nombres terminée!"
    
End Sub


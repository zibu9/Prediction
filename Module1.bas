Attribute VB_Name = "Module1"
Sub ConvertirChaineEnNombre()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim colonnes As Variant
    Dim valeurNettoyee As String

    ' D�finir la feuille de travail active
    Set ws = ThisWorkbook.Sheets("Soccer") ' Remplacez par le nom de votre feuille
    
    ' Colonnes � traiter
    colonnes = Array("J", "K", "L", "T", "U")
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne J (ou la plus longue des colonnes concern�es)
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Boucler � travers chaque ligne
    For i = 9 To lastRow ' Commence � la ligne 9 comme indiqu�
    
        ' Boucler � travers chaque colonne � traiter
        For j = LBound(colonnes) To UBound(colonnes)
            With ws.Cells(i, colonnes(j))
                ' Nettoyer la valeur : enlever les espaces et remplacer le point par une virgule
                valeurNettoyee = Trim(.Value) ' Enlever les espaces
                valeurNettoyee = Replace(valeurNettoyee, ".", ",") ' Remplacer le point par une virgule
                
                ' V�rifier si la cellule contient une cha�ne de caract�re num�rique et convertir
                If IsNumeric(valeurNettoyee) Then
                    .Value = CDbl(valeurNettoyee) ' Convertir en double
                End If
            End With
        Next j
        
    Next i
    
    MsgBox "Conversion des cha�nes de caract�res en nombres termin�e!"
    
End Sub


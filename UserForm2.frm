VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Tirage Aleatoire"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTirage_Click()

    Dim sheetName As String
    Dim lineCount As Integer

    ' Récupérer les valeurs saisies par l'utilisateur
    sheetName = txtSheetName.Value
    lineCount = CInt(txtLineCount.Value)

    ' Appeler le module de tirage aléatoire
    Call TirageAleatoire(sheetName, lineCount)

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Prediction"
   ClientHeight    =   1584
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3720
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call ConvertirChaineEnNombre
    Call CreerColonnePrediction
    Call CreerColonnePrediction2
    Call CreerColonnePrediction3
    Call CreerColonnePrediction4
    Call VerifierEtRemplirDecision
    Call ColorerLignes
    MsgBox "Les modules ont �t� ex�cut�s avec succ�s!"
End Sub

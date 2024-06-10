Attribute VB_Name = "Ribbon"
Option Explicit


'-------------------------------------------
' Au click du bouton de recherche
' suppression des valeurs dans la feuille
' ouverture du formulaire
'-------------------------------------------
Public Sub generateOneSearchNutrition(control As IRibbonControl)

Application.ScreenUpdating = False

With Ws_Nutrition
  .Range("B8:B32").Value = ""
  .Range("ASupprimer").ClearContents
  .Range("A1").Select
End With

With Ws_Nutrition.Shapes("InsertIMG").Fill
  .Visible = msoTrue
  .Solid
  .ForeColor.RGB = RGB(131, 204, 235)
End With
    
Application.ScreenUpdating = True

    USF_Search.Show
    
End Sub

'-------------------------------------------------
' Bouton de reinitialisation du ruban
' suppression des donnees dans les celulles
' suppression de l'image dans le shape
'-------------------------------------------------

Public Sub generateOneSuppressionFicheNutrition(control As IRibbonControl)

Application.ScreenUpdating = False

With Ws_Nutrition.Shapes("InsertIMG").Fill
  .Visible = msoTrue
  .Solid
  .ForeColor.RGB = RGB(131, 204, 235)
End With

With Ws_Nutrition
  .Range("B8:B32").Value = ""
  .Range("ASupprimer").ClearContents
  .Range("A1").Select
End With

Application.ScreenUpdating = True

End Sub

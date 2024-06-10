Attribute VB_Name = "Save_CodeEtName"
Option Explicit
Const ColonneDepart = 2
Const LigneDepart = 7

'-------------------------------------------
' Chargement de la listbox du formulaire
' Recherche d'un produit selon la saisie
' Retour dans sa totalite de l'objet
' recuperation du nom des produits
' et du code correspondant
'-------------------------------------------

Public Sub ChargementListBox_Codes()

Dim tempFilePath As String, _
    dic, _
    trouve As Boolean, _
    result As Object, _
    fichier As String, _
    retour As String, _
    count As Long, _
    listResultat

On Error GoTo erreur
DoEvents

 ' Initialisation du path du fichier temporaire
  tempFilePath = Environ$("temp") & "\" & "tempSaveCodeName.txt"
   
   'Initialisation du compteur
    count = 1
   
Application.ScreenUpdating = False

With USF_Search

   Set result = SendAPIMoteurRecherche(TraduireGoogle(.Txt_Saisie.Value, "Fr", "En"))
   
   If result Is Nothing Then Set result = Nothing: Exit Sub
   
   fichier = Dir(tempFilePath)
     
     If fichier <> "" Then Kill tempFilePath ' Supprimez le fichier temporaire si le fichier existe
 
 ' Enregistrement des codes et des noms des produits
   Open tempFilePath For Output As #1

   For Each dic In result
    trouve = True
    retour = count & "," & TraduireGoogle(dic("product_name"), "en", "fr") & "," & dic("code")
    Print #1, retour
    count = count + 1
   Next

   Close #1

   If trouve = False Then 'si aucune information trouvee
   
     .Cb_Search.Value = "  Aucune information trouvée "
     
   Else
      
        Open tempFilePath For Input As #1
   
        Do While Not EOF(1)
            
          Line Input #1, dic
          listResultat = Split(dic, ",")
          USF_Search.Cb_Search.AddItem listResultat(0) & "| " & listResultat(1)
             
        Loop
         
        Close #1
     
   End If
  
Set result = Nothing

End With

Application.ScreenUpdating = True
Exit Sub

erreur:

   MsgBox Err.Description
   On Error GoTo 0
   
End Sub


'-------------------------------------------
' Insertion des valeurs dans leurs zones
' respective selon la selection faites
' dans la combobox
'-------------------------------------------

Public Sub AffichageResultatRecherche(valeur As String)

Dim result, _
    tempFilePath As String, _
    fichier As String, _
    listResultat, _
    dic, _
    jsonObject As Object, _
    jsonObjectResult As Object, _
    i As Integer
   
   result = Split(valeur, "|")
   
   tempFilePath = Environ$("temp") & "\" & "tempSaveCodeName.txt"
   
   fichier = Dir(tempFilePath)
   
   If fichier = "" Then MsgBox "Aucun document trouvé", vbCritical: Exit Sub
   
    Open tempFilePath For Input As #1
   
        Do While Not EOF(1)
            
          Line Input #1, dic
          listResultat = Split(dic, ",")
          If listResultat(0) = result(0) Then result = listResultat(UBound(listResultat)): Exit Do
             
        Loop
         
    Close #1
      
Application.ScreenUpdating = False

'---------------------------------
' Insertion des informations
' recherchées
'---------------------------------
On Error Resume Next

    dic = SendAPICodeUnique(CStr(result))
    
    Set jsonObject = ParseJson(dic)("product")
    
 With Ws_Nutrition
 '-----------------------------------------------------
 '
 ' Affichage du resultat du nom du produit
 '
 '-----------------------------------------------------
                                        
     .Range("NomProduit") = listResultat(1)
     
 '-----------------------------------------------------
 '
 ' Affichage du resultat du nutriscore
 '
 '-----------------------------------------------------
   
    Set jsonObjectResult = jsonObject("nutriscore_2023_tags")

    If Err = 0 Then

       .Range("Nutriscore") = jsonObjectResult(1)

    ElseIf Err > 0 Then

       .Range("Nutriscore").Value = "Aucune informations trouvées"

    End If
    
On Error GoTo 0

 '-----------------------------------------------------
 '
 ' Affichage des ingredients recherches
 '
 '-----------------------------------------------------
On Error Resume Next

     result = jsonObject("ingredients_text_fr")
     
     If result <> "" Then
        
       result = Split(result, ",")
       
       For i = 0 To UBound(result)
    
           .Cells(LigneDepart + i + 1, ColonneDepart) = result(i)
                  
       Next i
       
     Else
     
        Set jsonObjectResult = jsonObject("nutriments")
       
            If jsonObjectResult.count <> 0 Then
        
                For i = 1 To jsonObjectResult.count

                     .Cells(LigneDepart + i, ColonneDepart) = jsonObjectResult(i)("text")

                Next i
                 
            ElseIf jsonObjectResult.count = 0 Then
                .Range("B8").Value = "Aucune information trouvée"
            End If
        
        Set jsonObjectResult = Nothing
        
     End If
     
On Error GoTo 0

 '-----------------------------------------------------
 '
 ' Affichage du resultat total et 100G par produit
 '
 '-----------------------------------------------------

    Set jsonObjectResult = jsonObject("nutriments")
    
    ' Section des valeurs pour la totalité du produit
    .Range("Glucide") = jsonObjectResult("carbohydrates")
    .Range("Graisse") = jsonObjectResult("fat")
    .Range("fibre") = jsonObjectResult("fiber")
    .Range("sucre") = jsonObjectResult("sugars")
    .Range("sel") = jsonObjectResult("sals")
    .Range("Energy") = jsonObjectResult("energy")
    .Range("Proteine") = jsonObjectResult("proteins")
    .Range("sodium") = jsonObjectResult("sodium")
    .Range("kcal") = jsonObjectResult("energy-kcal")
    .Range("kj") = jsonObjectResult("energy-kj")
    
    ' Section des valeurs pour 100g
    .Range("Glucide100") = jsonObjectResult("carbohydrates_100g")
    .Range("Graisse100") = jsonObjectResult("fat_100g")
    .Range("Fibre100") = jsonObjectResult("fiber_100g")
    .Range("sucre_100") = jsonObjectResult("sugars_100g")
    .Range("sels100") = jsonObjectResult("sals_100g")
    .Range("energy100g") = jsonObjectResult("energy_100g")
    .Range("Proteine100") = jsonObjectResult("proteins_100g")
    .Range("sodium100") = jsonObjectResult("sodium_100g")
    .Range("Kcal100") = jsonObjectResult("energy-kcal_100g")
    .Range("Kj_100") = jsonObjectResult("energy-kj_100g")
    

 End With
    
'    '---------------------------------
'    ' Insertion de la photo
'    '---------------------------------
On Error Resume Next

      Send_PhotoAPI (ParseJson(dic)("product")("image_url"))

On Error GoTo 0

Application.ScreenUpdating = True

End Sub

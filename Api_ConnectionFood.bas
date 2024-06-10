Attribute VB_Name = "Api_ConnectionFood"
Option Explicit

'-------------------------------------------------
' API Connection au site openfood fact
' Recherche d'un produit selon la saisie
' Effectu� dans le formulaire 'USF_Search'
'-------------------------------------------------

Public Function SendAPICodeUnique(code As String)

Dim httpRequest As Object, _
    url As String, _
    reponse As String
    
    url = "https://world.openfoodfacts.org/api/v2/product/code=" & code & "?lc=fr" 'url du site avec le code de recherche
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    With httpRequest
       
       .Open "GET", url, False
       
       On Error Resume Next
       .send
       On Error GoTo 0
    
       If .Status = 200 Then
           SendAPICodeUnique = .responseText
       Else
          MsgBox "Erreur n�" & .Status, vbCritical
       End If
       
    End With
    
Set httpRequest = Nothing

End Function

'-------------------------------------------
' API Connection au site openfood fact
' Recherche d'un produit selon la saisie
' Retour dans sa totalite de l'objet
'-------------------------------------------

Public Function SendAPIMoteurRecherche(search As String) As Object

Dim httpRequest As Object, _
    url As String, _
    reponse As String, _
    result As Object, _
    dic
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP") 'initialisation de l'objet HTTP

    url = "https://world.openfoodfacts.org/api/v2/search?categories_tags_en=" & search
    
    With httpRequest

       .Open "GET", url, False 'ouverture de la requ�te en mode lecture
       
       On Error Resume Next
       .send
       On Error GoTo 0
       
       If .Status = 200 Then
            Set result = JsonConverter.ParseJson(httpRequest.responseText)("products") 'recuperation des donnees de l'API avec l'aide du JSON
            Set SendAPIMoteurRecherche = result
       Else
         MsgBox "Erreur n�" & .Status, vbCritical
       End If
    End With

Set result = Nothing
Set httpRequest = Nothing

End Function

'------------------------------------------
' API Connection au site openfood fact
' Recherche de l'image dans l'api selon
' son code
' Insertion de l'image recherche dans
' le shape de la feuille
'------------------------------------------

Sub Send_PhotoAPI(code As String)

Dim imgURL As String, _
    img As Object, _
    shp As Shape, _
    http As Object, _
    tempFilePath As String, _
    oResp() As Byte
    
On Error GoTo Error

    ' URL de l'image que je souhaite inserer
     imgURL = code
     
    ' Initialisation du shape pour l'insertion  de l'image
    Set shp = Ws_Nutrition.Shapes("InsertIMG")

    ' Cr�ez un objet pour effectuer une requ�te HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
    
        .Open "GET", imgURL, False 'requ�te GET � l'API pour r�cup�rer l'image
        .send
        
        'insertion de l'image dans la variable
        oResp = .responseBody
        
        ' V�rifiez si la requ�te a r�ussi
        If .Status = 200 Then

         ' Enregistrez l'image temporairement
           tempFilePath = Environ$("temp") & "\" & "tempimage.jpg"

            Open tempFilePath For Binary As #1
            Put #1, 1, oResp
            Close #1

            ' Ins�rez l'image dans le shape sp�cifi�e

            With shp.Fill
             .Visible = msoTrue
             .UserPicture tempFilePath
            End With

            ' Supprimez le fichier temporaire
            Kill tempFilePath

        Else
           MsgBox "Erreur n�" & .Status, vbCritical
        End If
        
    End With

Set shp = Nothing
Set http = Nothing
Set img = Nothing

Exit Sub
Error:

   MsgBox Err.Description
   On Error GoTo 0
   Set shp = Nothing
   Set http = Nothing
   Set img = Nothing
   
End Sub



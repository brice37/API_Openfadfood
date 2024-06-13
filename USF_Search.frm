VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Search 
   Caption         =   " "
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
   OleObjectBlob   =   "USF_Search.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TousMesControls(1 To 2) As New Visualisation_USF

Private Sub Btn1_Click()
 
With Me.Txt_Saisie

  If .Value <> "" Then
    ChargementListBox_Codes
  Else
    .Value = "Aucune saisie effectuée"
    .ForeColor = vbRed
  End If
  
End With

End Sub

Private Sub Btn2_Click()

With Me.Cb_Search
  
  If .Value <> "" Then
     AffichageResultatRecherche (Me.Cb_Search.Value)
     Unload Me
  Else
    .Value = "Aucune selection effectuée"
    .ForeColor = vbRed
  End If
  
End With

End Sub

Private Sub Txt_Saisie_Enter()
 
 With Me
   .Cb_Search.Clear
   .Cb_Search.Value = ""
   .Txt_Saisie.Value = ""
 End With
 
End Sub


Private Sub UserForm_Initialize()

Me.Txt_Saisie.SetFocus

Dim i As Integer
   
   For i = 1 To 2
     Set TousMesControls(i).CMDUserform = Me("Btn" & i)
     Set TousMesControls(i).TxtUserform = Me("Txt_Saisie")
     Set TousMesControls(i).FrmUserform3 = Me("Frame3")
     Set TousMesControls(i).CBUserform = Me("Cb_Search")
     Set TousMesControls(i).LblUserform = Me("Lbl1")
     Set TousMesControls(i).FrmUserform = Me("Frame" & i)
   Next i
        
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Visualisation_USF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const Color1 = &H8000000F
Const Color2 = &H80000012
Const Color3 = vbBlack
Const Color4 = &HFF0000
Const Color5 = vbWhite

Public WithEvents CMDUserform As MSForms.CommandButton
Attribute CMDUserform.VB_VarHelpID = -1
Public WithEvents TxtUserform As MSForms.TextBox
Attribute TxtUserform.VB_VarHelpID = -1
Public WithEvents CBUserform As MSForms.ComboBox
Attribute CBUserform.VB_VarHelpID = -1
Public WithEvents FrmUserform As MSForms.Frame
Attribute FrmUserform.VB_VarHelpID = -1
Public WithEvents FrmUserform3 As MSForms.Frame
Attribute FrmUserform3.VB_VarHelpID = -1
Public WithEvents LblUserform As MSForms.Label
Attribute LblUserform.VB_VarHelpID = -1


Private Sub CBUserform_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  With CBUserform
    If .Value = "Aucune selection effectuée" Then
        .Value = ""
        .ForeColor = Color3
    End If
  End With

  Exit Sub

End Sub

Private Sub CMDUserform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim ctrl As control

   For Each ctrl In USF_Search.Controls

     If TypeOf ctrl Is MSForms.CommandButton Then

        With ctrl
           .BackColor = Color1
           .ForeColor = Color2
        End With

        With CMDUserform
           .BackColor = Color3
           .ForeColor = Color5
        End With

      End If
   Next
   
End Sub

Private Sub FrmUserform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

CommandButton_Couleur

With LblUserform
  .BackColor = Color1
  .ForeColor = Color4
End With

End Sub

Private Sub FrmUserform3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With LblUserform
  .BackColor = Color1
  .ForeColor = Color4

End With

End Sub

Private Sub LblUserform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With LblUserform
  .BackColor = Color4
  .ForeColor = Color1
End With

End Sub

Private Sub TxtUserform_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  With TxtUserform
    .Value = ""
    .ForeColor = Color3
  End With
  
  Exit Sub
  
End Sub

Private Sub TxtUserform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

CommandButton_Couleur
   
End Sub

Private Sub CBUserform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

CommandButton_Couleur
   
End Sub

Private Sub CommandButton_Couleur()

Dim ctrl As control
   
   For Each ctrl In USF_Search.Controls
     
     If TypeOf ctrl Is MSForms.CommandButton Then
        
        With CMDUserform
           .BackColor = Color1
           .ForeColor = Color2
        End With
        
     End If
     
   Next

End Sub

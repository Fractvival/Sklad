VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm20 
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9225
   OleObjectBlob   =   "UserForm20.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ThisWorkbook.changeText = "ANO"
    NoTopMost Me
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    ThisWorkbook.changeText = "NE"
    NoTopMost Me
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    SetTopMost Me
End Sub

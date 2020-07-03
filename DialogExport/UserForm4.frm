VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Chyba databáze "
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    NoTopMost Me
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    SetTopMost Me
    CommandButton1.SetFocus
End Sub

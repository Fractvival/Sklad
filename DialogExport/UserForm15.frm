VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm15 
   Caption         =   "REŽIM SPRÁVCE"
   ClientHeight    =   12510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16530
   OleObjectBlob   =   "UserForm15.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    NoTopMost Me
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    UserForm17.Show
End Sub

Private Sub Frame1_Click()

End Sub


Private Sub UserForm_Activate()
    HideBar Me
    SetTopMostAndFullscreen Me
    Frame1.Left = ((Me.Width - Me.Left) / 2) - (Frame1.Width / 2)
    Frame1.Top = ((Me.Height - Me.Top) / 2) - (Frame1.Height / 2)
    
End Sub


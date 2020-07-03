VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Úspìšné odepsání dílu"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm5"
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
    Label7.Caption = ThisWorkbook.GetMyKZM
    Label8.Caption = ThisWorkbook.GetMyPartNumber
    Label9.Caption = ThisWorkbook.GetMyName1 & ThisWorkbook.GetMyName2
    Label10.Caption = ThisWorkbook.searchText
    If (ThisWorkbook.changeBool) Then
        Label11.Caption = ThisWorkbook.changeText
        ThisWorkbook.changeBool = False
        ThisWorkbook.changeText = ""
    Else
        Label11.Caption = "1"
    End If
End Sub

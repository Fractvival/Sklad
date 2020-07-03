VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm10 
   Caption         =   "UserForm10"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14985
   OleObjectBlob   =   "UserForm10.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim textValue As Integer
    textValue = CInt(TextBox1.Text)
    Dim repoCount As Integer
    repoCount = CInt(Label12.Caption)
    repoCount = repoCount - textValue
    Dim writeLog As UserLog
    Set writeLog = New UserLog
    writeLog.writeLog 2, Label6.Caption, Label7.Caption, Label9.Caption, CStr(textValue), Label8.Caption, "ODPIS POMOCI HLEDANI"
    If SetNewCount(Label6.Caption, repoCount) = False Then
        SetNewCount Label7.Caption, repoCount
    End If
    Dim rCount As String
    Dim Name1 As String
    Dim Name2 As String
    rCount = CStr(repoCount)
    Name1 = ThisWorkbook.GetMyName1()
    Name2 = ThisWorkbook.GetMyName2()
    ThisWorkbook.SetMyInfo Label6.Caption, Label7.Caption, Name1, Name2, rCount, Label8.Caption
    NoTopMost Me
    Me.Hide
End Sub
Private Sub CommandButton2_Click()
    NoTopMost Me
    Me.Hide
End Sub
Private Sub CommandButton3_Click()
    Dim textValue As Integer
    textValue = CInt(TextBox1.Text)
    If (textValue <= 1) Then
        GoTo EndBtn
    End If
    textValue = textValue - 1
    TextBox1.Text = CStr(textValue)
EndBtn:
End Sub
Private Sub CommandButton4_Click()
    Dim textValue As Integer
    textValue = CInt(TextBox1.Text)
    If (textValue >= 99) Then
        GoTo EndBtn
    End If
    textValue = textValue + 1
    TextBox1.Text = CStr(textValue)
EndBtn:
End Sub
Private Sub UserForm_Activate()
    HideBar Me
    SetTopMost Me
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    GetINFO (ThisWorkbook.globalKZM)
    Label6.Caption = ThisWorkbook.GetMyKZM()
    Label7.Caption = ThisWorkbook.GetMyPartNumber()
    Label8.Caption = ThisWorkbook.GetMyRepo()
    Label9.Caption = ThisWorkbook.GetMyName1() & " " & ThisWorkbook.GetMyName2()
    Label12.Caption = ThisWorkbook.GetMyCount()
    TextBox1.Text = "1"
End Sub

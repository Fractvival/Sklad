VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm18 
   Caption         =   "UserForm18"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13725
   OleObjectBlob   =   "UserForm18.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    NumlockON
    NoTopMost Me
    Me.Hide
    ''pokud je pomocna promenna bool na true, vyzadujeme pouze cisla v textu
    If (ThisWorkbook.changeBool = True) Then
        If (IsNumeric(TextBox1.Text) = True) Then
            ThisWorkbook.searchText = TextBox1.Text
        Else
            UserForm1.Show
        End If
    Else
        ThisWorkbook.searchText = TextBox1.Text
    End If
End Sub

Private Sub CommandButton10_Click()
    ICT ("9")
End Sub

Private Sub CommandButton11_Click()
    ICT ("0")
End Sub

Private Sub CommandButton12_Click()
    ICT ("Q")
End Sub

Private Sub CommandButton13_Click()
    ICT ("W")
End Sub

Private Sub CommandButton14_Click()
    ICT ("E")
End Sub

Private Sub CommandButton15_Click()
    ICT ("R")
End Sub

Private Sub CommandButton16_Click()
    ICT ("T")
End Sub

Private Sub CommandButton17_Click()
    ICT ("Z")
End Sub

Private Sub CommandButton18_Click()
    ICT ("U")
End Sub

Private Sub CommandButton19_Click()
    ICT ("I")
End Sub

Private Sub CommandButton2_Click()
    ICT ("1")
End Sub

Private Sub CommandButton20_Click()
    ICT ("O")
End Sub

Private Sub CommandButton21_Click()
    ICT ("P")
End Sub

Private Sub CommandButton22_Click()
    ICT ("A")
End Sub

Private Sub CommandButton23_Click()
    ICT ("S")
End Sub

Private Sub CommandButton24_Click()
    ICT ("D")
End Sub

Private Sub CommandButton25_Click()
    ICT ("F")
End Sub

Private Sub CommandButton26_Click()
    ICT ("G")
End Sub

Private Sub CommandButton27_Click()
    ICT ("H")
End Sub

Private Sub CommandButton28_Click()
    ICT ("J")
End Sub

Private Sub CommandButton29_Click()
    ICT ("K")
End Sub

Private Sub CommandButton3_Click()
    ICT ("2")
End Sub

Private Sub CommandButton30_Click()
    ICT ("L")
End Sub

Private Sub CommandButton31_Click()
    ICT (":")
End Sub

Private Sub CommandButton32_Click()
    ICT ("Y")
End Sub

Private Sub CommandButton33_Click()
    ICT ("X")
End Sub

Private Sub CommandButton34_Click()
    ICT ("C")
End Sub

Private Sub CommandButton35_Click()
    ICT ("V")
End Sub

Private Sub CommandButton36_Click()
    ICT ("B")
End Sub

Private Sub CommandButton37_Click()
    ICT ("N")
End Sub

Private Sub CommandButton38_Click()
    ICT ("M")
End Sub

Private Sub CommandButton39_Click()
    ICT (".")
End Sub

Private Sub CommandButton4_Click()
    ICT ("3")
End Sub

Private Sub CommandButton40_Click()
    ICT ("-")
End Sub

Private Sub CommandButton41_Click()
    TextBox1.SetFocus
    SendKeys ("{BACKSPACE}")
End Sub

Private Sub CommandButton42_Click()
    ThisWorkbook.searchText = ""
    NoTopMost Me
    Me.Hide
End Sub

Private Sub CommandButton43_Click()
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub CommandButton44_Click()
    ICT (" ")
End Sub

Private Sub CommandButton45_Click()
    ICT ("�")
End Sub

Private Sub CommandButton46_Click()
    ICT ("�")
End Sub

Private Sub CommandButton47_Click()
    ICT ("�")
End Sub

Private Sub CommandButton48_Click()
    ICT ("�")
End Sub

Private Sub CommandButton49_Click()
    ICT ("�")
End Sub

Private Sub CommandButton5_Click()
    ICT ("4")
End Sub

Private Sub CommandButton50_Click()
    ICT ("�")
End Sub

Private Sub CommandButton51_Click()
    ICT ("�")
End Sub

Private Sub CommandButton52_Click()
    ICT ("�")
End Sub

Private Sub CommandButton53_Click()
    ICT ("�")
End Sub

Private Sub CommandButton54_Click()
    ICT ("�")
End Sub

Private Sub CommandButton55_Click()
    TextBox1.SetFocus
    SendKeys ("{LEFT}")
End Sub

Private Sub CommandButton56_Click()
    TextBox1.SetFocus
    SendKeys ("{RIGHT}")
End Sub

Private Sub CommandButton6_Click()
    ICT ("5")
End Sub

Private Sub CommandButton7_Click()
    ICT ("6")
End Sub

Private Sub CommandButton8_Click()
    ICT ("7")
End Sub

Private Sub CommandButton9_Click()
    ICT ("8")
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox1_Change()
    If (Len(TextBox1.Text) > 0) Then
        CommandButton1.enabled = True
    Else
        CommandButton1.enabled = False
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim lenNormal As Integer
    lenNormal = 35
    If (ThisWorkbook.changeNumber > 0) Then
        lenNormal = ThisWorkbook.changeNumber
    End If
    If (Len(TextBox1.Text)) >= lenNormal Then
        If (KeyCode <> 8) Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub ICT(InsCharText As String)
    Dim lenICT As Integer
    Dim lenNormal As Integer
    lenNormal = 36
    If (ThisWorkbook.changeNumber > 0) Then
        lenNormal = (ThisWorkbook.changeNumber + 1)
    End If
    lenICT = Len(InsCharText)
    If (Len(TextBox1.Text) < (lenNormal - lenICT)) Then
        Dim pos As Integer
        Dim save As String
        TextBox1.SetFocus
        pos = TextBox1.SelStart
        TextBox1.SelStart = pos
        TextBox1.SelLength = Len(TextBox1.Text)
        save = TextBox1.SelText
        TextBox1.Cut
        If (pos = 0) Then
            TextBox1.Text = InsCharText & save
        Else
            TextBox1.Text = TextBox1.Text & InsCharText & save
        End If
        TextBox1.SelStart = pos + lenICT
    End If
End Sub

Private Sub UserForm_Activate()
    HideBar Me
    SetTopMost Me
    Frame1.Left = (Me.Width / 2) - (Frame1.Width / 2)
    Frame1.Top = (Me.Height / 2) - (Frame1.Height / 2)
    TextBox1.Text = ""
    Label1.Caption = ThisWorkbook.searchText
    CommandButton1.enabled = False
    ThisWorkbook.searchText = ""
    If (ThisWorkbook.changeText <> "") Then
        TextBox1.Text = ThisWorkbook.changeText
    End If
    Label1.BackColor = RGB(220, 170, 50)
    Label1.BorderColor = RGB(255, 0, 0)
    NumlockON
    TextBox1.SetFocus
End Sub


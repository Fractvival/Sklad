VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Sklad"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12300
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''STIKNUTO tlacitko PRIHLASIT
Private Sub CommandButton12_Click()
    NumlockON
    ''TESTOVANI ZDA-LI je zadane heslo CISLO
    ''pokud se budou pouzivat jine nez ciselne hesla, muze se tato podminka odstranit
    If Not IsNumeric(TextBox1.Text) Then
        UserForm1.Show
    Else
        Dim i As Long
        Dim bCheck As Boolean
        bCheck = False
        ''ZDE hledame, jestli je heslo ulozeno v souboru uzivatelu
        For i = 1 To ThisWorkbook.GetTotalID()
            ''pokud heslo bude nalezeno v zaznamu, muzeme uzivatele prihlasit
            ''tj., ulozime jeho ID do pomocne promenne v ThisWorkbook
            If (TextBox1.Text = ThisWorkbook.GetPassFromID(ThisWorkbook.GetID(i))) Then
                ThisWorkbook.SetLoginID (ThisWorkbook.GetID(i))
                bCheck = True
            End If
        Next i
        ''HESLO je v poradku, muzeme dialog prihlaseni ukoncit
        If (bCheck = True) Then
            TextBox1.Text = ""
            NoTopMost Me
            Me.Hide
        Else
            UserForm6.Show
        End If
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton10_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 0
    End If
    TextBox1.SetFocus
End Sub


Private Sub CommandButton14_Click()
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub CommandButton15_Click()
    If (Len(TextBox1.Text) > 0) Then
        If (TextBox1.PasswordChar = "*") Then
            TextBox1.PasswordChar = ""
        Else
            TextBox1.PasswordChar = "*"
        End If
        CommandButton12.SetFocus
    End If
End Sub

Private Sub CommandButton9_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 9
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton8_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 8
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton7_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 7
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton6_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 6
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton5_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 5
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton4_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 4
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton3_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 3
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton2_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 2
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton1_Click()
    If Len(TextBox1.Text) < 6 Then
        TextBox1.Text = TextBox1.Text & 1
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton11_Click()
    If Len(TextBox1.Text) > 0 Then
        TextBox1.Text = Mid(TextBox1.Text, 1, Len(TextBox1.Text) - 1)
    End If
    TextBox1.SetFocus
End Sub
Private Sub CommandButton13_Click()
    Workbooks.Application.Quit
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox1_Change()
    CommandButton12.Default = True
    If (Len(TextBox1.Text) > 0) Then
        CommandButton12.enabled = True
    Else
        CommandButton12.enabled = False
    End If
End Sub
Private Sub UserForm_Activate()
    HideBar Me
    SetTopMostAndFullscreen Me
    CommandButton13.Left = (Me.Width - Me.Left) - 40
    Frame1.Left = ((Me.Width - Me.Left) / 2) - (Frame1.Width / 2)
    Frame1.Top = ((Me.Height - Me.Top) / 2) - (Frame1.Height / 2)
    TextBox1.Text = ""
    TextBox1.SetFocus
    ThisWorkbook.SetLoginID ("")
    CommandButton12.enabled = False
    NumlockON
    Frame1.SetFocus
    TextBox1.SetFocus
End Sub



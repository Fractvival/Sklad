VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm14 
   Caption         =   "UserForm14"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14280
   OleObjectBlob   =   "UserForm14.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim remCount As Integer
Dim repText As Boolean
Dim saveText As String


''tlacitko ODEPSAT
Private Sub CommandButton1_Click()
    ''Toto tlacitko neni videt (rozmery 0x0px), avsak je na nej nastaven fokus z duvodu "zastaveni" enteru (stisku)
    ''..to slouzi k potlaceni potvrzeni po hromadnem odectu ze skeneru
    ''nachazi se uplne nahore uprostred
    ''CommandButton49.SetFocus
    NumlockON
    ThisWorkbook.ResetMyInfo
    ''pomoci fce GIQ zjistime, zda-li se v databazi skladu nachazi PRESNA hodnota zadana k odepsani
    If (GetINFOQuick(saveText) = True) Then
        saveText = ""
        Dim oldNumber As Integer
        Dim newNumber As Integer
        Dim writeLog As UserLog
        Set writeLog = New UserLog ''pripravime k zapisu logu
        oldNumber = CInt(ThisWorkbook.GetMyCount) ''zjistime puvodni pocet dilu na skladu
        If (remCount > 1) Then
            newNumber = oldNumber - remCount ''..hromadny odecet (skener)
        Else
            newNumber = oldNumber - 1 ''..a odecteme jeden kus
        End If
        repText = False
        ThisWorkbook.searchText = CStr(newNumber) ''novy pocet si ulozime do pomocne promenne v ThisWorkbook
        ThisWorkbook.changeText = CStr(remCount) ''hromadny pocet do pomocne promenne v ThisWorkbook
        ThisWorkbook.changeBool = True
        ''PRESNOU hodnotu odepiseme ze skladu pomoci KZM
        If (SetNewCountQuick(ThisWorkbook.GetMyKZM, newNumber) = False) Then
            ''pokud KZM neni, tak ji ulozime pomoci PartNumber (cislo dilu)
            If (SetNewCountQuick(ThisWorkbook.GetMyPartNumber, newNumber) = False) Then
                ''pokud i presto nastane chyba, vyhodime hlasku
                NoTopMost Me
                Me.Hide
                UserForm3.Show
            Else
                ''v pripade uspechu zapiseme do logu a zobrazime hlasku o uspechu
                writeLog.writeLog 2, ThisWorkbook.GetMyKZM, ThisWorkbook.GetMyPartNumber, ThisWorkbook.GetMyName1 & " " & ThisWorkbook.GetMyName2, CStr(remCount), ThisWorkbook.GetMyRepo, "RYCHLY ODPIS - KZM"
                remCount = 0
                NoTopMost Me
                Me.Hide
                UserForm5.Show
            End If
        Else
            writeLog.writeLog 2, ThisWorkbook.GetMyKZM, ThisWorkbook.GetMyPartNumber, ThisWorkbook.GetMyName1 & " " & ThisWorkbook.GetMyName2, CStr(remCount), ThisWorkbook.GetMyRepo, "RYCHLY ODPIS - PN"
            remCount = 0
            NoTopMost Me
            Me.Hide
            UserForm5.Show
        End If
    Else
        NoTopMost Me
        Me.Hide
        UserForm3.Show
    End If
    remCount = 0
End Sub
''cislo 9
Private Sub CommandButton10_Click()
    ICT ("9")
End Sub
''cislo 0
Private Sub CommandButton11_Click()
    ICT ("0")
End Sub
''pismeno Q
Private Sub CommandButton12_Click()
    ICT ("Q")
End Sub
''Pismeno W
Private Sub CommandButton13_Click()
    ICT ("W")
End Sub
''pismeno E
Private Sub CommandButton14_Click()
    ICT ("E")
End Sub
''pismeno R
Private Sub CommandButton15_Click()
    ICT ("R")
End Sub
''pismeno T
Private Sub CommandButton16_Click()
    ICT ("T")
End Sub
''pismeno Z
Private Sub CommandButton17_Click()
    ICT ("Z")
End Sub
''pismeno U
Private Sub CommandButton18_Click()
    ICT ("U")
End Sub
''pismeno I
Private Sub CommandButton19_Click()
    ICT ("I")
End Sub
''cislo 1
Private Sub CommandButton2_Click()
    ICT ("1")
End Sub
''pismeno O
Private Sub CommandButton20_Click()
    ICT ("O")
End Sub
''pismeno P
Private Sub CommandButton21_Click()
    ICT ("P")
End Sub
''pismeno A
Private Sub CommandButton22_Click()
    ICT ("A")
End Sub
''pismeno S
Private Sub CommandButton23_Click()
    ICT ("S")
End Sub
''pismeno D
Private Sub CommandButton24_Click()
    ICT ("D")
End Sub
''pismeno F
Private Sub CommandButton25_Click()
    ICT ("F")
End Sub
''pismeno G
Private Sub CommandButton26_Click()
    ICT ("G")
End Sub
''pismeno H
Private Sub CommandButton27_Click()
    ICT ("H")
End Sub
''pismeno J
Private Sub CommandButton28_Click()
    ICT ("J")
End Sub
''pismeno K
Private Sub CommandButton29_Click()
    ICT ("K")
End Sub
''pismeno 2
Private Sub CommandButton3_Click()
    ICT ("2")
End Sub
''pismeno L
Private Sub CommandButton30_Click()
    ICT ("L")
End Sub
''znak :
Private Sub CommandButton31_Click()
    ICT (":")
End Sub
''pismeno Y
Private Sub CommandButton32_Click()
    ICT ("Y")
End Sub
''pismeno X
Private Sub CommandButton33_Click()
    ICT ("X")
End Sub
''pismeno C
Private Sub CommandButton34_Click()
    ICT ("C")
End Sub
''pismeno V
Private Sub CommandButton35_Click()
    ICT ("V")
End Sub
''pismeno B
Private Sub CommandButton36_Click()
    ICT ("B")
End Sub
''pismeno N
Private Sub CommandButton37_Click()
    ICT ("N")
End Sub
''pismeno M
Private Sub CommandButton38_Click()
    ICT ("M")
End Sub
''znak .
Private Sub CommandButton39_Click()
    ICT (".")
End Sub
''cislo 3
Private Sub CommandButton4_Click()
    ICT ("3")
End Sub
''znak -
Private Sub CommandButton40_Click()
    ICT ("-")
End Sub
''backspace
Private Sub CommandButton41_Click()
    TextBox1.SetFocus
    SendKeys ("{BACKSPACE}")
End Sub
''zavreni dialogu krizkem
Private Sub CommandButton42_Click()
    ThisWorkbook.searchText = ""
    NoTopMost Me
    Me.Hide
End Sub
''vymazani celeho textboxu
Private Sub CommandButton43_Click()
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub
''mezernik
Private Sub CommandButton44_Click()
    ICT (" ")
End Sub
''320.
Private Sub CommandButton45_Click()
    If Len(TextBox1.Text) < 32 Then
        ICT ("320.")
    End If
End Sub
''32.
Private Sub CommandButton46_Click()
    If Len(TextBox1.Text) < 33 Then
        ICT ("32.")
    End If
End Sub

Private Sub CommandButton47_Click()
    TextBox1.SetFocus
    SendKeys ("{LEFT}")
End Sub

Private Sub CommandButton48_Click()
    TextBox1.SetFocus
    SendKeys ("{RIGHT}")
End Sub

''cislo 4
Private Sub CommandButton5_Click()
    ICT ("4")
End Sub
''cislo 5
Private Sub CommandButton6_Click()
    ICT ("5")
End Sub
''cislo 6
Private Sub CommandButton7_Click()
    ICT ("6")
End Sub
''cislo 7
Private Sub CommandButton8_Click()
    ICT ("7")
End Sub
''cislo 8
Private Sub CommandButton9_Click()
    ICT ("8")
End Sub

Private Sub Frame1_Click()

End Sub

Sub Timer(Finish As Long)
    Dim NowTick As Long
    Dim EndTick As Long
    EndTick = GetTickCount + (Finish)
    Do
        NowTick = GetTickCount
        DoEvents
    Loop Until NowTick >= EndTick
End Sub


''zmena textu v textboxu
Private Sub TextBox1_Change()
    If (Len(TextBox1.Text) > 0) Then
        CommandButton1.enabled = True ''pokud je v textboxu text, zpristupni tlacitko Odepsat
    Else
        CommandButton1.enabled = False
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If (Len(TextBox1.Text)) >= 35 Then
        If (KeyCode <> 8) Then
            KeyCode = 0
        End If
    End If
    
    '' ZDE, mechanizmus pro spocitani kusu k odpisu pomoci skeneru (hromadny odecet)
    '' JE NUTNE, aby TextBox mel polozku EnterKeyBehaviour nastaveno na True
    If (KeyCode = 13) Then '' pokud je striknut enter
        repText = True
        saveText = TextBox1.Text
        Dim i As Integer
        Do While repText
            If (TextBox1.Text <> "") Then
                remCount = remCount + 1
                TextBox1.Text = ""
                i = 0
            End If
            Timer (25)
            i = i + 1
            If (i > 20) Then
                repText = False
                CommandButton1.enabled = True
            End If
        Loop
    End If
    
End Sub

Private Sub ICT(InsCharText As String)
    Dim lenICT As Integer
    lenICT = Len(InsCharText)
    If (Len(TextBox1.Text) < (36 - lenICT)) Then
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

''po aktivaci (zobrazeni) dialogu
Private Sub UserForm_Activate()
    HideBar Me
    SetTopMost Me
    TextBox1.Text = ""
    TextBox1.SetFocus
    CommandButton1.enabled = False
    ThisWorkbook.searchText = ""
End Sub


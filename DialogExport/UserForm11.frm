VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "UserForm11"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15645
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentDate As Date ''pomocna promenna aktualni datum
''nacteni/vypsani zaznamu dle mesice a roku do listobxu
Private Sub LoadToList()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    Dim valMonth As Integer
    Dim valYear As Integer
    valMonth = 1
    valYear = CInt(TextBox3.Text)
    Select Case UCase(TextBox2.Text)
        Case "LEDEN"
            valMonth = 1
        Case "ÚNOR"
            valMonth = 2
        Case "BØEZEN"
            valMonth = 3
        Case "DUBEN"
            valMonth = 4
        Case "KVÌTEN"
            valMonth = 5
        Case "ÈERVEN"
            valMonth = 6
        Case "ÈERVENEC"
            valMonth = 7
        Case "SRPEN"
            valMonth = 8
        Case "ZÁØÍ"
            valMonth = 9
        Case "ØÍJEN"
            valMonth = 10
        Case "LISTOPAD"
            valMonth = 11
        Case "PROSINEC"
            valMonth = 12
    End Select
    If Len(Dir(ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & valYear & ".xlsx", vbNormal)) = 0 Then
        GoTo NormalHandler
    End If
    Workbooks.Open Filename:=ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & valYear, ReadOnly:=True
    ActiveWorkbook.Sheets(valMonth).Select
    Dim i As Integer
    Dim lastRow As Integer
    Dim Text As Date
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
    Text = Cells(i, 1)
    If (CInt(Month(Text)) = valMonth) Then
        If (Cells(i, 3) = "ODEBRAT") Then
            ListBox1.AddItem Format(Cells(i, 1), "Short Date") & " | " & Format(Cells(i, 2), "Long Time") & " | " & Format(Cells(i, 4), "General Number") & " | " & Format(Cells(i, 5), "General Number") & " | " & Cells(i, 6) & " | " & Format(Cells(i, 7), "General Number") & " | " & Cells(i, 8), -1
        End If
    End If
    Next i
    ActiveWorkbook.Close
    If (ListBox1.ListCount > 0) Then
        ListBox1.SetFocus
        ListBox1.ListIndex = 0
        CommandButton1.enabled = True
    End If
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
NormalHandler:
End Sub
''Stiknuto tlacitko ODEBRAT ZAZNAM logu
Private Sub CommandButton1_Click()
    ''vypneme update "obrazovky" a hlaseni
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    Dim valMonth As Integer
    Dim valYear As Integer
    valMonth = 1 ''pocatecni hodnota mesice pri spusteni dialogu (Leden)
    valYear = CInt(TextBox3.Text) ''Prevod roku na cislo
    ''Nyni prevody mesicu z textu na cisla
    Select Case UCase(TextBox2.Text)
        Case "LEDEN"
            valMonth = 1
        Case "ÚNOR"
            valMonth = 2
        Case "BØEZEN"
            valMonth = 3
        Case "DUBEN"
            valMonth = 4
        Case "KVÌTEN"
            valMonth = 5
        Case "ÈERVEN"
            valMonth = 6
        Case "ÈERVENEC"
            valMonth = 7
        Case "SRPEN"
            valMonth = 8
        Case "ZÁØÍ"
            valMonth = 9
        Case "ØÍJEN"
            valMonth = 10
        Case "LISTOPAD"
            valMonth = 11
        Case "PROSINEC"
            valMonth = 12
    End Select
    Dim KZM As String
    Dim PartNumber As String
    Dim Nazev As String
    Dim Pocet As String
    Dim Misto As String
    Dim PuvodniDatum As String
    Dim PuvodniCas As String
    ''Pokud neexistuje soubor s logy, neni co vypsat do listboxu, z toho duvodu ukoncime dialog
    If Len(Dir(ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & valYear & ".xlsx", vbNormal)) = 0 Then
        GoTo NormalHandler ''presun na NormalHandler, tedy ukonceni dialogu
    End If
    ''otevreme soubor s logy uzivatele pro zvoleny mesic a rok (soubor je otevren pro zapis)
    Workbooks.Open Filename:=ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & valYear, ReadOnly:=False
    ''vybereme list daneho mesice a nastavime jako aktivni (zvoleny)
    ActiveWorkbook.Sheets(valMonth).Select
    Dim i As Integer
    Dim lastRow As Integer
    Dim Text As Date
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row ''zjistime pocet radku zaznamu
    Dim selList As Integer
    Dim itemCount As Integer
    selList = ListBox1.ListIndex ''zde si ulozime aktualne zvolenou polozku v listboxu
    itemCount = 0 ''pomocnou promennou pro pocitani polozek nastavime na 0
    On Error GoTo ErrHandler ''pri chybe pri praci se soubory se presun do ErrorHandleru
    ''nyni, zacneme prohledavat radek po radku vsechny zaznamy, zazname typu ODEBRAT nas zajimaji..
    ''Pokud se pomocna promenna (zvysovana v tomto cyklu for..next) bude rovnat vybrane polozce listboxu..
    ''..tak zapocneme s odebiranim zaznamu
    ''Pocitani zacne od radku 2, prvni radek jsou popisy sloupcu
    For i = 2 To lastRow
    Text = Cells(i, 1)
    If (CInt(Month(Text)) = valMonth) Then
        If (Cells(i, 3) = "ODEBRAT") Then ''pokud sedi typ zaznamu
            If (itemCount = selList) Then ''a pokud je shoda pomocne promenne a zvolene polozky v listoboxu..
                ''ulozime si do pomocnych promennych informace o zaznamu urceneho k smazani
                KZM = Format(Cells(i, 4), "General Number")
                PartNumber = Format(Cells(i, 5), "General Number")
                Nazev = Cells(i, 6)
                Pocet = Format(Cells(i, 7), "General Number")
                Misto = Cells(i, 8)
                PuvodniDatum = Format(Cells(i, 1), "Long Date")
                PuvodniCas = Format(Cells(i, 2), "Long Time")
                ListBox1.RemoveItem (selList) ''..vymaz z listobxu zvoleny zaznam
                Rows(i).EntireRow.Delete ''a taktez vymaz zaznam z logu v souboru
            End If
            itemCount = itemCount + 1
        End If
    End If
    Next i
    ActiveWorkbook.Close saveChanges:=True ''soubor s logy ulozime
    ''zde resim ktere polozka, jestli vubec nejaka, bude v listboxu po smazani zvolena jako aktivni
    If (ListBox1.ListCount >= 1) Then
        ListBox1.SetFocus
    End If
    If (ListBox1.ListCount = 0) Then
        ListBox1.SetFocus
        CommandButton1.enabled = False
    End If
    
    Dim writeLog As UserLog
    Set writeLog = New UserLog
    ''Nyni, zapiseme novy log o tom, ze doslo ke smazani zaznamu v souboru logu
    writeLog.writeLog 3, KZM, PartNumber, Nazev, Pocet, Misto, "(" & PuvodniDatum & " " & PuvodniCas & ")"
    ''A TED, je potreba vratit stav poctu dilu ve skladu na hodnotu, ktera byla pred odebranim dilu
    If (KZM <> "") Then
        AddNewCount KZM, CInt(Pocet)
    Else
        AddNewCount PartNumber, CInt(Pocet)
    End If
    On Error GoTo 0
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    GoTo NormalHandler
ErrHandler:
    UserForm13.Show
NormalHandler:
End Sub
''TLACITKO ZMENIT DATUM - Zmena datumu dle prani uzivatele a vypsani novych logu
Private Sub CommandButton3_Click()
    ListBox1.Clear
    LoadToList
End Sub

Private Sub Frame1_Click()

End Sub

' MECHANIZMUS PRI POSUNU SPIN BUTTONEM ROK DOLU
Private Sub SpinButton4_SpinDown()
    If (TextBox3.Text <> "") Then
        TextBox3.Text = SpinButton4.Value
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM ROK NAHORU
Private Sub SpinButton4_SpinUp()
    If (TextBox3.Text <> "") Then
        TextBox3.Text = SpinButton4.Value
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM MESIC DOLU
Private Sub SpinButton3_SpinDown()
    If (TextBox2.Text <> "") Then
        TextBox2.Text = UCase(MonthName(SpinButton3.Value, False))
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM MESIC NAHORU
Private Sub SpinButton3_SpinUp()
    If (TextBox2.Text <> "") Then
        TextBox2.Text = UCase(MonthName(SpinButton3.Value, False))
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM DEN DOLU
Private Sub SpinButton2_SpinDown()
    If (TextBox1.Text <> "") Then
        TextBox1.Text = SpinButton2.Value
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM DEN NAHORU
Private Sub SpinButton2_SpinUp()
    If (TextBox1.Text <> "") Then
        TextBox1.Text = SpinButton2.Value
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM1 DOLU
Private Sub SpinButton1_SpinDown()
    If (ListBox1.ListCount > 1) Then
        If (ListBox1.ListIndex < (ListBox1.ListCount - 1)) Then
            ListBox1.ListIndex = (ListBox1.ListIndex + 1)
        End If
    End If
End Sub
' MECHANIZMUS PRI POSUNU SPIN BUTTONEM1 NAHORU
Private Sub SpinButton1_SpinUp()
    If (ListBox1.ListCount > 1) Then
        If (ListBox1.ListIndex >= 1) Then
            Dim oldPos As Long
            oldPos = ListBox1.ListIndex
            ListBox1.ListIndex = (ListBox1.ListIndex - 1)
        End If
    End If
End Sub
''zavreni tohoto dialogu pomoci krizku
Private Sub CommandButton2_Click()
    NoTopMost Me
    Me.Hide
End Sub
''zobrazeni tohoto dialogu
Private Sub UserForm_Activate()
    HideBar Me
    SetTopMost Me
    currentDate = Now ''zjistime a ulozime aktualni datum a cas
    CommandButton1.enabled = False ''tlacitko pro odebrani zneplatnime
    TextBox3.Text = year(currentDate)
    TextBox2.Text = UCase(MonthName(Month(currentDate), False))
    TextBox1.Text = Day(currentDate)
    ''nastaveni rozsahu spinbuttonu
    SpinButton2.min = 1
    SpinButton2.Max = 31
    SpinButton2.SmallChange = 1
    SpinButton2.Value = Day(currentDate)
    SpinButton3.min = 1
    SpinButton3.Max = 12
    SpinButton3.SmallChange = 1
    SpinButton3.Value = Month(currentDate)
    SpinButton4.min = 2020
    SpinButton4.Max = 2100
    SpinButton4.SmallChange = 1
    SpinButton4.Value = year(currentDate)
    ''vymazeme vse v listboxu
    ListBox1.Clear
    ''a nacteme do listoboxu logy pro aktualni mesic a rok
    LoadToList
End Sub

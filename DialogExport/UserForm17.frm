VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm17 
   Caption         =   "Editor uživatelù"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13950
   OleObjectBlob   =   "UserForm17.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totalID As Long
Dim name As String
Dim id As String
Dim pass As String
Dim su As String
Dim other As String

Private Function writeUser() As Boolean
    writeUser = True
    ''vypneme update "obrazovky" a hlaseni
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    On Error GoTo ErrHandler ''pri chybe pri praci se soubory se presun do ErrorHandleru
    Workbooks.Open Filename:=ThisWorkbook.pathData & ThisWorkbook.FilenameUsers, ReadOnly:=False
    ActiveWorkbook.Sheets("UZIVATEL").Visible = -1
    ActiveWorkbook.Sheets("UZIVATEL").Select
    Dim lastRow As Integer
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row ''zjistime pocet radku zaznamu
    Range("A" & lastRow + 1).Value = pass
    Range("B" & lastRow + 1).Value = id
    Range("C" & lastRow + 1).Value = name
    Range("D" & lastRow + 1).Value = su
    Range("E" & lastRow + 1).Value = other
    ''ListBox1.AddItem id & " | " & name & " | " & pass & " | " & su & " | " & other, -1
    ActiveWorkbook.Sheets("UZIVATEL").Visible = 0
    ActiveWorkbook.Close saveChanges:=True
    ''zde resim ktere polozka, jestli vubec nejaka, bude v listboxu po smazani zvolena jako aktivni
    On Error GoTo 0
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    GoTo NormalHandler
ErrHandler:
    writeUser = False
    UserForm13.Show
NormalHandler:
End Function

Private Function writeExistUser(userIndex As Long) As Boolean
    writeExistUser = True
    ''vypneme update "obrazovky" a hlaseni
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    On Error GoTo ErrHandler ''pri chybe pri praci se soubory se presun do ErrorHandleru
    Workbooks.Open Filename:=ThisWorkbook.pathData & ThisWorkbook.FilenameUsers, ReadOnly:=False
    ActiveWorkbook.Sheets("UZIVATEL").Visible = -1
    ActiveWorkbook.Sheets("UZIVATEL").Select
    Dim lastRow As Integer
    ''lastRow = Cells(Rows.Count, 1).End(xlUp).Row ''zjistime pocet radku zaznamu
    Range("A" & userIndex + 1).Value = pass
    ''Range("B" & userIndex + 1).Value = id
    Range("C" & userIndex + 1).Value = name
    Range("D" & userIndex + 1).Value = su
    Range("E" & userIndex + 1).Value = other
    ''ListBox1.AddItem id & " | " & name & " | " & pass & " | " & su & " | " & other, -1
    ActiveWorkbook.Sheets("UZIVATEL").Visible = 0
    ActiveWorkbook.Close saveChanges:=True
    ''zde resim ktere polozka, jestli vubec nejaka, bude v listboxu po smazani zvolena jako aktivni
    On Error GoTo 0
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    GoTo NormalHandler
ErrHandler:
    writeExistUser = False
    UserForm13.Show
NormalHandler:
End Function

Private Function removeUser(userIndex As Long) As Boolean
    removeUser = True
    ''vypneme update "obrazovky" a hlaseni
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    On Error GoTo ErrHandler ''pri chybe pri praci se soubory se presun do ErrorHandleru
    Workbooks.Open Filename:=ThisWorkbook.pathData & ThisWorkbook.FilenameUsers, ReadOnly:=False
    ActiveWorkbook.Sheets("UZIVATEL").Visible = -1
    ActiveWorkbook.Sheets("UZIVATEL").Select
    Dim lastRow As Integer
    ''lastRow = Cells(Rows.Count, 1).End(xlUp).Row ''zjistime pocet radku zaznamu
    Range("A" & userIndex + 1).EntireRow.Delete
    ActiveWorkbook.Sheets("UZIVATEL").Visible = 0
    ActiveWorkbook.Close saveChanges:=True
    ''zde resim ktere polozka, jestli vubec nejaka, bude v listboxu po smazani zvolena jako aktivni
    On Error GoTo 0
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    GoTo NormalHandler
ErrHandler:
    removeUser = False
    UserForm13.Show
NormalHandler:
End Function


Private Sub CommandButton1_Click()
    name = ""
    id = ""
    pass = ""
    su = ""
    other = ""
    ThisWorkbook.searchText = "KROK 1/5:  ZADEJ JEDINEÈNÉ ÈÍSELNÉ ID (MAX. 6 ÈÍSEL)"
    ThisWorkbook.changeNumber = 6
    ThisWorkbook.changeBool = True
    UserForm18.Show
    id = ThisWorkbook.searchText
    If id <> "" Then
        If (ThisWorkbook.GetPassFromID(id) = "") Then
            ThisWorkbook.changeNumber = 35
            ThisWorkbook.changeBool = False
            ThisWorkbook.searchText = "KROK 2/5:  ZADEJ JMÉNO A PØIJMENÍ (MAX. 35 ZNAKÙ)"
            UserForm18.Show
            name = ThisWorkbook.searchText
            If name <> "" Then
                ThisWorkbook.changeNumber = 6
                ThisWorkbook.changeBool = True
                ThisWorkbook.searchText = "KROK 3/5:  ZADEJ ÈÍSELNÉ HESLO (MAX.6 ÈÍSEL)"
                UserForm18.Show
                pass = ThisWorkbook.searchText
                If pass <> "" Then
                    ThisWorkbook.changeNumber = 1
                    ThisWorkbook.changeBool = True
                    ThisWorkbook.searchText = "KROK 4/5:  ZADEJ OPRÁVNÌNÍ: 0=UŽIVATEL, 1=ADMIN"
                    UserForm18.Show
                    su = ThisWorkbook.searchText
                    If su <> "" Then
                        ThisWorkbook.changeNumber = 35
                        ThisWorkbook.changeBool = False
                        ThisWorkbook.searchText = "KROK 5/5:  ZADEJ DODATKY (MAX. 35 ZNAKÙ)"
                        UserForm18.Show
                        other = ThisWorkbook.searchText
                        If other <> "" Then
                            '' MAME VSE CO POTREBUJEME
                            ''MsgBox name & " | " & id & " | " & pass & " | " & su & " | " & other
                            writeUser
                            lgnID = ThisWorkbook.GetLoginID ''ulozime id aktualne prihlaseneho spravce
                            ''refreshneme uzivatele
                            ThisWorkbook.LoadUsers
                            ThisWorkbook.SetLoginID (lgnID) ''a zpatky "prihlasime" spravce :]
                            refreshList
                        '' *
                        End If ''other <> ""
                    End If ''su <> ""
                End If ''pass <> ""
            End If ''name <> ""
        Else ''getpassid = "" neni splneno, cili pokud uz toto id ma nastaveno heslo..
            UserForm19.Show ''...zobrazime dialog o duplicite
        End If
    End If ''id <> ""

End Sub

Private Sub CommandButton2_Click()
    ThisWorkbook.changeText = ""
    
    Dim selItem As Long
    Dim lgnID As String
    
    selItem = ListBox1.ListIndex + 1
    
    Dim idd As String
    Dim oldSu As String
    idd = ThisWorkbook.GetID(selItem)
    id = idd
    name = ThisWorkbook.GetNameFromID(id)
    pass = ThisWorkbook.GetPassFromID(id)
    su = ThisWorkbook.GetSU(id)
    other = ThisWorkbook.GetOtherFromID(id)
    
    ThisWorkbook.searchText = "KROK 1/4:  ZADEJ JMÉNO A PØIJMENÍ (MAX. 35 ZNAKÙ)"
    ThisWorkbook.changeNumber = 35
    ThisWorkbook.changeBool = False
    ThisWorkbook.changeText = name
    UserForm18.Show
    name = ThisWorkbook.searchText
    If name <> "" Then
        ThisWorkbook.searchText = "KROK 2/4:  ZADEJ ÈÍSELNÉ HESLO (MAX.6 ÈÍSEL)"
        ThisWorkbook.changeNumber = 6
        ThisWorkbook.changeBool = True
        ThisWorkbook.changeText = ThisWorkbook.GetPassFromID(idd)
        UserForm18.Show
        pass = ThisWorkbook.searchText
        If pass <> "" Then
            oldSu = su
            ThisWorkbook.searchText = "KROK 3/4:  ZADEJ OPRÁVNÌNÍ: 0=UŽIVATEL, 1=ADMIN"
            ThisWorkbook.changeNumber = 1
            ThisWorkbook.changeBool = True
            ThisWorkbook.changeText = ThisWorkbook.GetSU(idd)
            UserForm18.Show
            su = ThisWorkbook.searchText
            If su <> "" Then
                ThisWorkbook.searchText = "KROK 4/4:  ZADEJ DODATKY (MAX. 35 ZNAKÙ)"
                ThisWorkbook.changeNumber = 35
                ThisWorkbook.changeBool = False
                ThisWorkbook.changeText = ThisWorkbook.GetOtherFromID(idd)
                UserForm18.Show
                other = ThisWorkbook.searchText
                If other <> "" Then
                    '' MAME VSE CO POTREBUJEME
                    If (oldSu <> su) Then
                        If (idd = ThisWorkbook.GetLoginID) Then
                            UserForm20.Show
                            If (ThisWorkbook.changeText = "ANO") Then
                                ''ulozime zmeny v opravneni a ukoncime program
                                ListBox1.List(ListBox1.ListIndex, 0) = id & " | " & name & " | " & pass & " | " & su & " | " & other
                                writeExistUser (selItem)
                                ActiveWorkbook.Close saveChanges:=False
                                Application.Quit
                            Else
                            ''NIC se neulozi
                            End If
                        Else
                            ''ulozit zmeny bez zmeny opravneni
                            ListBox1.List(ListBox1.ListIndex, 0) = id & " | " & name & " | " & pass & " | " & su & " | " & other
                            writeExistUser (selItem)
                            lgnID = ThisWorkbook.GetLoginID ''ulozime id aktualne prihlaseneho spravce
                            ''refreshneme uzivatele
                            ThisWorkbook.LoadUsers
                            ThisWorkbook.SetLoginID (lgnID) ''a zpatky "prihlasime" spravce :]
                            refreshList
                        End If
                    Else
                        ''ulozit zmeny bez zmeny opravneni
                        ListBox1.List(ListBox1.ListIndex, 0) = id & " | " & name & " | " & pass & " | " & su & " | " & other
                        writeExistUser (selItem)
                        lgnID = ThisWorkbook.GetLoginID ''ulozime id aktualne prihlaseneho spravce
                        ''refreshneme uzivatele
                        ThisWorkbook.LoadUsers
                        ThisWorkbook.SetLoginID (lgnID) ''a zpatky "prihlasime" spravce :]
                        refreshList
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub refreshList()
    totalID = ThisWorkbook.GetTotalID
    ListBox1.Clear
    Dim id As String
    Dim name As String
    Dim pass As String
    Dim su As String
    Dim other As String
    Dim i As Long
    For i = 1 To totalID
        id = ThisWorkbook.GetID(i)
        name = ThisWorkbook.GetNameFromID(id)
        pass = ThisWorkbook.GetPassFromID(id)
        su = ThisWorkbook.GetSU(id)
        other = ThisWorkbook.GetOtherFromID(id)
        ListBox1.AddItem id & " | " & name & " | " & pass & " | " & su & " | " & other, i - 1
    Next i
    ListBox1.ListIndex = 0
    If (ListBox1.ListCount > 1) Then
        CommandButton3.enabled = True
    Else
        CommandButton3.enabled = False
    End If
End Sub

Private Sub CommandButton3_Click()
    
    Dim selItem As Long
    selItem = ListBox1.ListIndex + 1
    Dim idd As String
    idd = ThisWorkbook.GetID(selItem)
    
    If (idd = ThisWorkbook.GetLoginID) Then
        UserForm21.Show
    Else
        UserForm22.Show
        If (ThisWorkbook.changeText = "ANO") Then
            removeUser (selItem)
            ListBox1.RemoveItem (ListBox1.ListIndex)
            lgnID = ThisWorkbook.GetLoginID ''ulozime id aktualne prihlaseneho spravce
            ''refreshneme uzivatele
            ThisWorkbook.LoadUsers
            ThisWorkbook.SetLoginID (lgnID) ''a zpatky "prihlasime" spravce :]
            refreshList
        End If
    End If

End Sub

Private Sub CommandButton4_Click()
    NoTopMost Me
    ListBox1.Clear
    ThisWorkbook.searchText = ""
    ThisWorkbook.changeText = ""
    Me.Hide
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Activate()
    
    SetTopMost Me
    HideBar Me
    refreshList

End Sub

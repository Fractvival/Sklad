VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''//////////////////////////////////////////////
''
'' SKLAD v1.0
''
''//////////////////////////////////////////////
Option Explicit ' Nove promenne pouze pres "Dim"
''CESTY K SOUBORUM A SLOZKAM JSOU ULOZENY V LIST1 (ten je defaultne schovany)
''HODNOTY ZDE PRIMO PROGRAMOVE NENASTAVUJEME, NASTAVUJI SE PRAVE V LISTU "LIST1"
''PROGRAM HODNOTY Z LISTU NACTE A PRIRADI JE DO KONKRETNICH PROMENNYCH SAM
' Cesta k databazi bez zpetneho lomitka, napr.: "C:\Databaze"
Public pathData As String
' Cesta k zaloham databaze, napr.: "\Zalohy"
Public pathDataBackups As String
' Cesta k logum uzivatelu, napr.: "\Logy"
Public pathDataLogs As String
' Nazev souboru s uzivateli, napr.: "\Uzivatel.xlsx"
Public FilenameUsers As String
' Nazev souboru se skladem, napr.: "\Sklad.xlsx"
Public FilenameRepos As String
Dim sFolder As String ' Pomocna promenna pro dialog vyberu slozky
Dim sEdit As String ' Pomocna promenna pro text EditBoxu dialogu prihlaseni
Dim wbUsers As Workbook ' Sesit s uzivateli
Dim wbLogs As Workbook ' Sesit s logy
'+++++++++++[ databaze uzivatelu typu POLE ]+++++++++++
Dim HESLO(1 To 1000) As String
Dim id(1 To 1000) As String
Dim JMENO(1 To 1000) As String
Dim OPRAVNENI(1 To 1000) As String
Dim Dodatky(1 To 1000) As String
'+++++++++++++[ pomocne promenne pri ziskavani dat o dilu ]+++++++++++++++++++++++
Public globalKZM As String
Dim myKZM As String
Dim myPartNumber As String
Dim myName1 As String
Dim myName2 As String
Dim myCount As String
Dim myRepo As String
'++++++++++++++++++++++++++++++++++++
Dim iTotalID As Long ' Celkovy pocet uzivatelu v databazi
Dim loginID As String ' ID (osobni cislo) aktualne prihlaseneho uzivatele
Public searchText As String ' Text ziskany z kreslene klavesnice + obcasna pomocna promenna pro vymenu textu
Public changeText As String ' pomocna promenna pro vymenu textu mezi dialogy
Public changeNumber As Integer ''pomocna promenna pro vymenu cisel mezi dialogy
Public changeBool As Boolean ''pomocna promenna bool mezi dialogy

Public Function GetMyKZM() As String
    GetMyKZM = myKZM
End Function
Public Function GetMyPartNumber() As String
    GetMyPartNumber = myPartNumber
End Function
Public Function GetMyName1() As String
    GetMyName1 = myName1
End Function
Public Function GetMyName2() As String
    GetMyName2 = myName2
End Function
Public Function GetMyCount() As String
    If (Len(myCount) = 0) Then
        myCount = "0"
    End If
    GetMyCount = myCount
End Function
Public Function GetMyRepo() As String
    GetMyRepo = myRepo
End Function
Public Sub ResetMyInfo()
    myKZM = ""
    myPartNumber = ""
    myName1 = ""
    myName2 = ""
    myCount = ""
    myRepo = ""
End Sub
Public Sub SetMyInfo(KZM As String, PN As String, Name1 As String, Name2 As String, Count As String, Repo As String)
    myKZM = KZM
    myPartNumber = PN
    myName1 = Name1
    myName2 = Name2
    myCount = Count
    myRepo = Repo
End Sub
' Zjisti heslo uzivatele pomoci ID
Public Function GetPassFromID(sID As String) As String
    Dim i As Long
    Dim sPass As String
    sPass = ""
    For i = 1 To iTotalID
        If (sID = id(i)) Then
            sPass = HESLO(i)
        End If
    Next i
    GetPassFromID = sPass
End Function
' Zjisti dodatky uzivatele pomoci ID
Public Function GetOtherFromID(sID As String) As String
    Dim i As Long
    Dim sOther As String
    sOther = ""
    For i = 1 To iTotalID
        If (sID = id(i)) Then
            sOther = Dodatky(i)
        End If
    Next i
    GetOtherFromID = sOther
End Function
' Zjisti jmeno uzivatele pomoci ID
Public Function GetNameFromID(sID As String) As String
    Dim i As Long
    Dim sName As String
    sName = ""
    For i = 1 To iTotalID
        If (sID = id(i)) Then
            sName = JMENO(i)
        End If
    Next i
    GetNameFromID = sName
End Function
' Vrati ID aktualne prihlaseneho uzivatele
Public Function GetLoginID() As String
    GetLoginID = loginID
End Function
' Nastavi ID aktualne prihlaseneho uzivatele
Public Sub SetLoginID(Number As String)
    loginID = Number
End Sub
' Vrati ID uzivatele z nactenych uzivatelu
' Parametr je cislo 1 az POCET UZIVATELU (iTotalID)
Public Function GetID(Number As Long) As String
    GetID = id(Number)
End Function
' Vrati pocet uzivatelu nactenych z databaze
Public Function GetTotalID() As Long
    GetTotalID = iTotalID
End Function
' Vrati opravneni uzivatele pomoci ID
' Typicky, 0=bez opravneni
Public Function GetSU(sID As String) As String
    GetSU = "0"
    Dim i As Long
    For i = 1 To iTotalID
        If (sID = id(i)) Then
            GetSU = OPRAVNENI(i)
        End If
    Next i
End Function
' Vrati cestu k databazim
Public Function GetPathData() As String
    GetPathData = pathData
End Function
' Nacte uzivatele z databaze
' Cesta k databazi se nachazi v miste ulozenem ve skrytem List1
' Tuto cestu ulozi samotny ThisWorkbook po spusteni
Sub LoadUsers()
    On Error GoTo ErrHandler ' pri nejakych chybach skoc do ErrHandleru
    Application.ScreenUpdating = False ' nebude rusit zmenou aktivniho sesitu nebo listu
    Set wbUsers = Workbooks.Open(pathData & FilenameUsers, , True)
    ''nyni zjistime pocet uzivatelu v databazi a ulozime do promenne "iTotalID"
    
    wbUsers.Worksheets("UZIVATEL").Visible = -1
    wbUsers.Worksheets("UZIVATEL").Select
    iTotalID = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    wbUsers.Worksheets("UZIVATEL").Visible = 0
    
    ''iTotalID = wbUsers.Worksheets("UZIVATEL").Range("B2:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
    
    ''zjistujeme pocet radku v celem sloupci, radek 1 je vsak popis, proto odecteme jednicku
    iTotalID = iTotalID - 1
    ' a do pomocnych promennych typu POLE ulozime jejich informace
    Dim i As Long
    For i = 1 To 1000
        HESLO(i) = ""
        id(i) = ""
        JMENO(i) = ""
        OPRAVNENI(i) = ""
        Dodatky(i) = ""
    Next i
    For i = 1 To iTotalID
        HESLO(i) = wbUsers.Worksheets("UZIVATEL").Range("A" & (i + 1)).Value
        id(i) = wbUsers.Worksheets("UZIVATEL").Range("B" & (i + 1)).Value
        JMENO(i) = wbUsers.Worksheets("UZIVATEL").Range("C" & (i + 1)).Value
        OPRAVNENI(i) = wbUsers.Worksheets("UZIVATEL").Range("D" & (i + 1)).Value
        Dodatky(i) = wbUsers.Worksheets("UZIVATEL").Range("E" & (i + 1)).Value
    Next i
    ' VSE hotovo, zavreme sesit s uzivateli
    wbUsers.Close saveChanges:=False
    ' Jelikoz jsme dosli az sem, nic zvlastniho se nestalo a skocime do bloku NoErr
    GoTo NoErr
ErrHandler:
    ' CHYBA pri cteni informaci o uzivatelich, ukonci program..
    UserForm4.Show
    Workbooks.Application.Quit
    Workbooks.Close
NoErr:
    Application.ScreenUpdating = True ' opet povolime zobrazovani zmen sesitu a listu
End Sub
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' SAMOTNY PROGRAM - spousti se pouze pri otevreni celeho sesitu
Private Sub Workbook_Open()
    ' Cesta k databazi ze sesitu List1 a bunky B1
    '' TATO hodnota je pote pouzivana v ramci celeho programu (sesitu)
    ''Pokud chceme videt list1, je potreba nastavit hodnotu Visible na True ve vlastnostech listu
    pathData = List1.Range("B1").Value
    ''A dale nastavime ostatni cesty
    ' Cesta k zaloham databaze, napr.: "\Zalohy"
    pathDataBackups = List1.Range("B2").Value
    ' Cesta k logum uzivatelu, napr.: "\Logy"
    pathDataLogs = List1.Range("B3").Value
    ' Nazev souboru s uzivateli, napr.: "\Uzivatel.xlsx"
    FilenameUsers = List1.Range("B4").Value
    ' Nazev souboru se skladem, napr.: "\Sklad.xlsx"
    FilenameRepos = List1.Range("B5").Value
    ' Nyni otestujeme zdali cesta k databazi existuje
    ' Pokud delka textu z vystupu fce Dir je nula,
    ' spustime proces pro vyber nove slozky
    Application.DisplayAlerts = False
    If (Len(Dir(pathData, vbDirectory)) = 0) Then
        UserForm16.Show
    End If
    ''** Nyni otestujeme a pripadne vytvorime vsechny dalsi potrebne slozky a soubory
    ''V prvni rade zkusime, jestli jde na cilove misto ukladat data...
    If (Len(Dir(pathData & "\TestDIR", vbDirectory)) = 0) Then
        MkDir (pathData & "\TestDIR")
            If (Len(Dir(pathData & "\TestDIR", vbDirectory)) = 0) Then
                MsgBox "Lituji, ale na cílové místo nelze ukládat!"
                Workbooks.Application.Quit ' ..ukoncime cely proces
                Workbooks.Close
            Else
                RmDir (pathData & "\TestDIR")
            End If
    End If
    ''Test / vytvoreni slozky pro zalohy
    If (Len(Dir(pathData & pathDataBackups, vbDirectory)) = 0) Then
        MkDir (pathData & pathDataBackups)
    End If
    ''Test / vytvoreni slozky pro logy
    If (Len(Dir(pathData & pathDataLogs, vbDirectory)) = 0) Then
        MkDir (pathData & pathDataLogs)
    End If
    ''Test / vytvoreni souboru s uzivateli
    If (Len(Dir(pathData & FilenameUsers, vbNormal)) = 0) Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Dim wb As Workbook
        Set wb = Workbooks.Add
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "UZIVATEL"
        Dim forUsers As Integer
            For forUsers = 1 To 5
            Cells(1, forUsers).Select
            Selection.NumberFormat = "@"
            Selection.Interior.ColorIndex = 39
            Selection.RowHeight = 20
            Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
            Selection.Borders(xlEdgeBottom).Weight = xlThick
            Selection.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
            Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            Selection.Borders(xlEdgeLeft).Weight = xlThin
            Selection.Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
            Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            Selection.Borders(xlEdgeRight).Weight = xlThin
            Selection.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
            Select Case forUsers
                Case 1
                    Selection.ColumnWidth = 15
                    Selection.Value = "HESLO"
                Case 2
                    Selection.ColumnWidth = 15
                    Selection.Value = "ID"
                Case 3
                    Selection.ColumnWidth = 25
                    Selection.Value = "JMENO"
                Case 4
                    Selection.ColumnWidth = 15
                    Selection.Value = "OPRAVNENI"
                Case 5
                    Selection.ColumnWidth = 25
                    Selection.Value = "DODATKY"
            End Select
        Next forUsers
        Range("A:A").HorizontalAlignment = xlCenter
        Range("A:A").VerticalAlignment = xlCenter
        Range("A:A").NumberFormat = "@"
        Range("B:B").HorizontalAlignment = xlCenter
        Range("B:B").VerticalAlignment = xlCenter
        Range("B:B").NumberFormat = "@"
        Range("C:C").HorizontalAlignment = xlCenter
        Range("C:C").VerticalAlignment = xlCenter
        Range("C:C").NumberFormat = "@"
        Range("D:D").HorizontalAlignment = xlCenter
        Range("D:D").VerticalAlignment = xlCenter
        Range("D:D").NumberFormat = "@"
        Range("E:E").HorizontalAlignment = xlCenter
        Range("E:E").VerticalAlignment = xlCenter
        Range("E:E").NumberFormat = "@"
        Range("A2") = "123456"
        Range("B2") = "123456"
        Range("C2") = "Administrator"
        Range("D2") = "1"
        Range("E2") = "PRO PRVNÍ PØIHLÁŠENÍ!"
        wb.Sheets(3).Delete
        wb.Sheets(2).Delete
        wb.Sheets(2).Visible = False
        wb.Sheets(1).Select
        Range("A1").ColumnWidth = 35
        Range("A1").HorizontalAlignment = xlCenter
        Range("A1").VerticalAlignment = xlCenter
        Range("A1").NumberFormat = "@"
        Range("A1").RowHeight = 25
        Range("A1") = "UŽIVATELÉ verze 1.0"
        wb.SaveAs pathData & FilenameUsers
        wb.Close
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End If
    ''Test pritomnosti souboru se skladem
    ''OCEKAVAJI SE V NEM NASLEDUJICI SLOUPCE A JEJICH NAZVY
    '' A = "KZM"
    '' B = "Part Number"
    '' C = "Nazev1"
    '' D = "Nazev2"
    '' E = "Poèet"
    '' F = "Inventura skladu"
    '' G = "Umístìní"
    '' H = "Doplnìno dne"
    '' Sloupec F a sloupec H neni v programu pouzit, ale pocita se s timto poctem sloupcu a jejich umisteni.
    If (Len(Dir(pathData & FilenameRepos, vbNormal)) = 0) Then
        MsgBox "Ve složce s databází se nenachází soubor se skladem (napr.: Sklad.xlsx)"
        Workbooks.Application.Quit
        Workbooks.Close
    End If
    Application.DisplayAlerts = True
Start:
    LoadUsers ' nacteni uzivatelu z databaze
    ' pokud je pocet uzivatelu nenulovy, spusti program
    If (iTotalID > 0) Then
        UserForm2.Show ' zobrazi dialog pro login uzivatele
        Dim testLog As UserLog
        Set testLog = New UserLog
        If (testLog.SetClass = False) Then
            UserForm12.Show
            Workbooks.Application.Quit
            Workbooks.Close
        End If
        Dim writeLog As UserLog
        Set writeLog = New UserLog
        writeLog.writeLog 1, "", "", "", "", "", ""
        UserForm7.Show ' zobrazi hlavni dialog pro prihlaseneho uzivatele
        Unload UserForm1
        Unload UserForm2
        Unload UserForm3
        Unload UserForm4
        Unload UserForm5
        Unload UserForm6
        Unload UserForm7
        Unload UserForm8
        Unload UserForm9
        Unload UserForm10
        Unload UserForm11
        Unload UserForm12
        Unload UserForm13
        Unload UserForm14
        Unload UserForm15
        Unload UserForm16
        Unload UserForm17
        Unload UserForm18
        Unload UserForm19
        Unload UserForm20
        Unload UserForm21
        Unload UserForm22
        GoTo Start
    Else
        ' ..v opacnem pripade program ukonci
        Workbooks.Application.Quit
        Workbooks.Close
    End If
End Sub ' KONEC PROGRAMU


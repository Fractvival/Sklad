Attribute VB_Name = "Module1"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, _
                                                              ByVal dwFlags As LongPtr, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Function GetKeyboardState Lib "user32.dll" (ByVal lpKeyState As LongPtr) As Boolean
    Private Declare PtrSafe Function apiGetTickCount Lib "Kernel32" _
                Alias "QueryPerformanceCounter" _
                (cyTickCount As Currency) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32" _
                (ByVal Index As Long) As Long
    Public Declare PtrSafe Function SetForegroundWindow _
                         Lib "user32" _
                       (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal hWndInsertAfter As Long, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal cx As Long, _
                    ByVal cy As Long, _
                    ByVal wFlags As Long) As Long
#Else
    Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, _
                                                      ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Function GetKeyboardState Lib "user32.dll" (ByVal lpKeyState As Long) As Boolean
    Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
    Public Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
    Public Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" _
               (ByVal hwnd As Long) As Long
    Public Declare Function GetSystemMetrics Lib "user32" _
                (ByVal Index As Long) As Long
    Public Declare Function SetForegroundWindow _
                         Lib "user32" _
                       (ByVal hwnd As Long) As Long
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal hWndInsertAfter As Long, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal cx As Long, _
                    ByVal cy As Long, _
                    ByVal wFlags As Long) As Long
#End If

Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_NUMLOCK As Byte = &H90
Private Const NumLockScanCode As Byte = &H45

Public Sub ToggleNumlock(enabled As Boolean)
    Dim keystate(255) As Byte
    'Test current keyboard state.
    GetKeyboardState (VarPtr(keystate(0)))
    If (Not keystate(VK_NUMLOCK) And enabled) Or (keystate(VK_NUMLOCK) And Not enabled) Then
        'Send a keydown
        keybd_event VK_NUMLOCK, NumLockScanCode, KEYEVENTF_EXTENDEDKEY, 0&
        'Send a keyup
        keybd_event VK_NUMLOCK, NumLockScanCode, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0&
    End If
End Sub

Public Sub NumlockON()
    Dim keystate(255) As Byte
    GetKeyboardState (VarPtr(keystate(0)))
    If (keystate(VK_NUMLOCK) = False) Then
        keybd_event VK_NUMLOCK, NumLockScanCode, KEYEVENTF_EXTENDEDKEY, 0&
        keybd_event VK_NUMLOCK, NumLockScanCode, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0&
    End If
End Sub

' Schova titulkovy pruh okna
Sub HideBar(frm As Object)
    Dim Style As Long, Menu As Long, hWndForm As Long
    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
    Style = GetWindowLong(hWndForm, &HFFF0)
    Style = Style And Not &HC00000
    SetWindowLong hWndForm, &HFFF0, Style
    DrawMenuBar hWndForm
End Sub
' Nastavi okno do top-popredi a zobrazi ve fullscreenu
Sub SetTopMostAndFullscreen(frm As Object)
    Dim hWndForm As Long
    'Const SWP_NOMOVE = &H2
    'Const SWP_NOSIZE = &H1
    Const HWND_TOPMOST = -1
    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
    SetWindowPos hWndForm, HWND_TOPMOST, 0, 0, GetSystemMetrics(0), GetSystemMetrics(1), 0
End Sub
' Nastavi okno do top-popredi
Sub SetTopMost(frm As Object)
    Dim hWndForm As Long
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
    SetWindowPos hWndForm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub
' Odstrani oknu top-popredi
Sub NoTopMost(frm As Object)
    Dim hWndForm As Long
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const HWND_NOTOPMOST = -2
    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
    SetWindowPos hWndForm, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
' pomoci teto funkce dojde k vyhledani a nacteni informaci o nahradnim dilu
' Hledani probiha na zaklade KZM cisla a nebo na zaklade Part NUMBER
' Nalezeny dil a informace o nem jsou pomoci teto procedury ulozeny v ThisWorkbook..
' ..ulozeny jsou do promennych s prefixem "my[NAZEV]"
' tato fce rovnez resetuje predchozi ulozeni (bylo-li nejake) a to az po..
' ..kladnem nalezeni dilu...
Function GetINFO(NumberKZMorPart As String) As Boolean
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    Dim isFirstSearch As Boolean ' Detekce prvni nalezene polozky
    Dim isF As Boolean 'signalizuje uspesne nalezeni polozky (urceno pro vystup)
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    '+++++++++++++++++++++++
    'kontrolu uspechu nastavime na false
    isF = False
    'az najdeme prvni polozku, nastavime na true a vsechny dalsi polozky uz se vynechaji
    isFirstSearch = False
    'nyni otevreme sesit se skladem
    Set wb = Workbooks.Open(path, True, True)
    'nechceme aby nam sesit se skladem prekreslil nas otevreny sesit (tento), proto tuto akci schovame a
    'zrovna vypneme i zobrazovani chybovych hlasek atp...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'nyni nastavime text k vyhledavani
    'ten je predany jako parametr teto funkci pomoci promenne NumberKZMorPart
    'je to hodnota, kterou jsme ziskali z listboxu v dialogu pro vyhledavani
    'a kterou jsme ulozili do promenne v ThisWorkbook
    MyArr = Array(NumberKZMorPart)
    With wb.Sheets(1).Range("A2:B10000") 'hledat budeme jen mezi KZM a Part Number
        For i = LBound(MyArr) To UBound(MyArr)

            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlFormulas, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do
                    KZM = wb.Sheets(1).Range("A" & Rng.Row).Value
                    PartNumber = wb.Sheets(1).Range("B" & Rng.Row).Value
                    Name1 = wb.Sheets(1).Range("C" & Rng.Row).Value
                    Name2 = wb.Sheets(1).Range("D" & Rng.Row).Value
                    Count = wb.Sheets(1).Range("E" & Rng.Row).Value
                    Repo = wb.Sheets(1).Range("G" & Rng.Row).Value
                    'prvni nalezena polozka je nas cil, ostatni preskocime
                    If (isFirstSearch = False) Then
                        ThisWorkbook.ResetMyInfo
                        ThisWorkbook.SetMyInfo KZM, PartNumber, Name1, Name2, Count, Repo
                        isF = True
                        isFirstSearch = True
                    End If
                    Set Rng = .FindNext(Rng)
                Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
            End If
        Next i
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    wb.Close
    GetINFO = isF
End Function
Function GetINFOQuick(NumberKZMorPart As String) As Boolean
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    Dim isFirstSearch As Boolean ' Detekce prvni nalezene polozky
    Dim isF As Boolean 'signalizuje uspesne nalezeni polozky (urceno pro vystup)
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    '+++++++++++++++++++++++
    'kontrolu uspechu nastavime na false
    isF = False
    'az najdeme prvni polozku, nastavime na true a vsechny dalsi polozky uz se vynechaji
    isFirstSearch = False
    'nyni otevreme sesit se skladem
    Set wb = Workbooks.Open(path, True, True)
    'nechceme aby nam sesit se skladem prekreslil nas otevreny sesit (tento), proto tuto akci schovame a
    'zrovna vypneme i zobrazovani chybovych hlasek atp...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'nyni nastavime text k vyhledavani
    'ten je predany jako parametr teto funkci pomoci promenne NumberKZMorPart
    'je to hodnota, kterou jsme ziskali z listboxu v dialogu pro vyhledavani
    'a kterou jsme ulozili do promenne v ThisWorkbook
    MyArr = Array(NumberKZMorPart)
    With wb.Sheets(1).Range("A2:B10000") 'hledat budeme jen mezi KZM a Part Number
        For i = LBound(MyArr) To UBound(MyArr)

            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do
                    KZM = wb.Sheets(1).Range("A" & Rng.Row).Value
                    PartNumber = wb.Sheets(1).Range("B" & Rng.Row).Value
                    Name1 = wb.Sheets(1).Range("C" & Rng.Row).Value
                    Name2 = wb.Sheets(1).Range("D" & Rng.Row).Value
                    Count = wb.Sheets(1).Range("E" & Rng.Row).Value
                    Repo = wb.Sheets(1).Range("G" & Rng.Row).Value
                    'prvni nalezena polozka je nas cil, ostatni preskocime
                    If (isFirstSearch = False) Then
                        ThisWorkbook.ResetMyInfo
                        ThisWorkbook.SetMyInfo KZM, PartNumber, Name1, Name2, Count, Repo
                        isF = True
                        isFirstSearch = True
                    End If
                    Set Rng = .FindNext(Rng)
                Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
            End If
        Next i
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    wb.Close
    GetINFOQuick = isF
End Function
Function SetNewCount(NumberKZMorPart As String, newCount As Integer) As Boolean
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim isFirstSearch As Boolean ' Detekce prvni nalezene polozky
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    Dim isF As Boolean 'signalizuje uspesne nalezeni polozky (urceno pro vystup)
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    '+++++++++++++++++++++++
    'kontrolu uspechu nastavime na false
    isF = False
    'az najdeme prvni polozku, nastavime na true a vsechny dalsi polozky uz se vynechaji
    isFirstSearch = False
    'nechceme aby nam sesit se skladem prekreslil nas otevreny sesit (tento), proto tuto akci schovame a
    'zrovna vypne i zobrazovani chybovych hlasek atp...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'nyni otevreme sesit se skladem
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=False)
    MyArr = Array(NumberKZMorPart)
    Dim dateFormat As String
    dateFormat = Format(Now(), "yyyy_mm_dd_hh_mm_ss") 'zde jsme vyformatovali datum a cas pro nazev zalozniho souboru skladu
    If Len(Dir(ThisWorkbook.pathData & ThisWorkbook.pathDataBackups, vbDirectory)) = 0 Then
       MkDir ThisWorkbook.pathData & ThisWorkbook.pathDataBackups
    End If
    'pred zapisem do skladu provedeme zalohu, vysledny soubor (zaloha) bude obsahovat datum, cas a id cislo uzivatele
    ActiveWorkbook.SaveCopyAs ThisWorkbook.pathData & ThisWorkbook.pathDataBackups & "\" & dateFormat & "_" & ThisWorkbook.GetLoginID & ".xlsx"
    With wb.Sheets(1).Range("A2:B10000") 'opet hledame mezi KZM a PartNumber
        For i = LBound(MyArr) To UBound(MyArr)
            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlFormulas, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do
                    'prvni nalezena polozka je nas cil, ostatni preskocime
                    If (isFirstSearch = False) Then
                        'A ZDE provedeme zapis
                        wb.Sheets(1).Range("E" & Rng.Row).NumberFormat = "@"
                        wb.Sheets(1).Range("E" & Rng.Row) = newCount
                        isF = True
                        isFirstSearch = True
                    End If
                    Set Rng = .FindNext(Rng)
                Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
            End If
        Next i
    End With
    'Ulozime zmeny do skladu a soubor zavreme
    ActiveWorkbook.Close saveChanges:=True
    SetNewCount = isF 'vlajka uspechu
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Function
Function SetNewCountQuick(NumberKZMorPart As String, newCount As Integer) As Boolean
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim isFirstSearch As Boolean ' Detekce prvni nalezene polozky
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    Dim isF As Boolean 'signalizuje uspesne nalezeni polozky (urceno pro vystup)
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    '+++++++++++++++++++++++
    'kontrolu uspechu nastavime na false
    isF = False
    'az najdeme prvni polozku, nastavime na true a vsechny dalsi polozky uz se vynechaji
    isFirstSearch = False
    'nechceme aby nam sesit se skladem prekreslil nas otevreny sesit (tento), proto tuto akci schovame a
    'zrovna vypne i zobrazovani chybovych hlasek atp...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'nyni otevreme sesit se skladem
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=False)
    MyArr = Array(NumberKZMorPart)
    Dim dateFormat As String
    dateFormat = Format(Now(), "yyyy_mm_dd_hh_mm_ss") 'zde jsme vyformatovali datum a cas pro nazev zalozniho souboru skladu
    If Len(Dir(ThisWorkbook.pathData & ThisWorkbook.pathDataBackups, vbDirectory)) = 0 Then
       MkDir ThisWorkbook.pathData & ThisWorkbook.pathDataBackups
    End If
    'pred zapisem do skladu provedeme zalohu, vysledny soubor (zaloha) bude obsahovat datum, cas a id cislo uzivatele
    ActiveWorkbook.SaveCopyAs ThisWorkbook.pathData & ThisWorkbook.pathDataBackups & "\" & dateFormat & "_" & ThisWorkbook.GetLoginID & ".xlsx"
    With wb.Sheets(1).Range("A2:B10000") 'opet hledame mezi KZM a PartNumber
        For i = LBound(MyArr) To UBound(MyArr)
            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do
                    'prvni nalezena polozka je nas cil, ostatni preskocime
                    If (isFirstSearch = False) Then
                        'A ZDE provedeme zapis
                        wb.Sheets(1).Range("E" & Rng.Row).NumberFormat = "@"
                        wb.Sheets(1).Range("E" & Rng.Row) = newCount
                        isF = True
                        isFirstSearch = True
                    End If
                    Set Rng = .FindNext(Rng)
                Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
            End If
        Next i
    End With
    'Ulozime zmeny do skladu a soubor zavreme
    ActiveWorkbook.Close saveChanges:=True
    SetNewCountQuick = isF 'vlajka uspechu
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Function
Function AddNewCount(NumberKZMorPart As String, addCount As Integer) As Boolean
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim isFirstSearch As Boolean ' Detekce prvni nalezene polozky
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    Dim isF As Boolean 'signalizuje uspesne nalezeni polozky (urceno pro vystup)
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    '+++++++++++++++++++++++
    'kontrolu uspechu nastavime na false
    isF = False
    'az najdeme prvni polozku, nastavime na true a vsechny dalsi polozky uz se vynechaji
    isFirstSearch = False
    'nechceme aby nam sesit se skladem prekreslil nas otevreny sesit (tento), proto tuto akci schovame a
    'zrovna vypne i zobrazovani chybovych hlasek atp...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'nyni otevreme sesit se skladem
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=False)
    MyArr = Array(NumberKZMorPart)
    Dim dateFormat As String
    dateFormat = Format(Now(), "yyyy_mm_dd_hh_mm_ss") 'zde jsme vyformatovali datum a cas pro nazev zalozniho souboru skladu
    If Len(Dir(ThisWorkbook.pathData & ThisWorkbook.pathDataBackups & "\", vbDirectory)) = 0 Then
       MkDir ThisWorkbook.pathData & ThisWorkbook.pathDataBackups
    End If
    'pred zapisem do skladu provedeme zalohu, vysledny soubor (zaloha) bude obsahovat datum, cas a id cislo uzivatele
    ActiveWorkbook.SaveCopyAs ThisWorkbook.pathData & ThisWorkbook.pathDataBackups & "\" & dateFormat & "_" & ThisWorkbook.GetLoginID & ".xlsx"
    With wb.Sheets(1).Range("A2:B10000") 'opet hledame mezi KZM a PartNumber
        For i = LBound(MyArr) To UBound(MyArr)
            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlFormulas, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do
                    'prvni nalezena polozka je nas cil, ostatni preskocime
                    If (isFirstSearch = False) Then
                        'A ZDE provedeme zapis
                        Dim oldCount As String
                        oldCount = wb.Sheets(1).Range("E" & Rng.Row)
                        addCount = addCount + CInt(oldCount)
                        wb.Sheets(1).Range("E" & Rng.Row).NumberFormat = "@"
                        wb.Sheets(1).Range("E" & Rng.Row) = addCount
                        isF = True
                        isFirstSearch = True
                    End If
                    Set Rng = .FindNext(Rng)
                Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
            End If
        Next i
    End With
    'Ulozime zmeny do skladu a soubor zavreme
    ActiveWorkbook.Close saveChanges:=True
    AddNewCount = isF 'vlajka uspechu
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Function

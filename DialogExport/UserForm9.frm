VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "Výsledky hledání"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15915
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
'' VYHLEDAVANI
'' JADRO 1.0
'' PRACOVNI VERZE: 3.6.2020

Dim arrKZM(1 To 10000) As String

' ODEBRAT ZE SKLADU
Private Sub CommandButton1_Click()
    If (ListBox1.ListCount > 0) Then
        ThisWorkbook.globalKZM = ""
        ThisWorkbook.globalKZM = arrKZM(ListBox1.ListIndex + 1)
        ''NoTopMost Me
        ''Me.Hide
        UserForm10.Show
        
        Dim KZM As String
        Dim PartNumber As String
        Dim Name1 As String
        Dim Name2 As String
        Dim Count As String
        Dim Repo As String
        Dim nLine As String
        
        KZM = ThisWorkbook.GetMyKZM()
        PartNumber = ThisWorkbook.GetMyPartNumber()
        Name1 = ThisWorkbook.GetMyName1()
        Name2 = ThisWorkbook.GetMyName2()
        Count = ThisWorkbook.GetMyCount()
        Repo = ThisWorkbook.GetMyRepo()
        
        If (KZM <> "") Then nLine = "KZM: " & KZM
        If (PartNumber <> "") Then nLine = nLine & " | ID: " & PartNumber
        If (Name1 <> "") Then nLine = nLine & " | Nazev: " & Name1
        If (Name2 <> "") Then nLine = nLine & " " & Name2
        If (Count <> "") Then nLine = nLine & " | Pocet: " & Count
        If (Repo <> "") Then nLine = nLine & " | Misto: " & Repo
        
        Dim selList As Integer
        selList = ListBox1.ListIndex
        ListBox1.List(selList) = nLine
        ''ListBox1.RemoveItem selList
        ''ListBox1.AddItem nLine, selList
        
    Else
        Application.Speech.Speak "Please, find new parts."
    End If
End Sub

' NOVE HLEDANI
Private Sub CommandButton2_Click()
    NoTopMost Me
    Me.Hide
    UserForm8.Show
    If (ThisWorkbook.searchText <> "") Then
        ListBox1.Clear
        Label2.Caption = ""
        rCount = 0
        nLine = ""
    End If
    Me.Show
End Sub
' UZAVRENI DIALOGU VYHLEDAVANI KRIZKEM
Private Sub CommandButton5_Click()
    ListBox1.Clear
    Label2.Caption = ""
    rCount = 0
    nLine = ""
    NoTopMost Me
    Me.Hide
End Sub
' PRI ZMENE POLOZKY V LISTBOXU
Private Sub ListBox1_Change()
    Label3.Caption = (ListBox1.ListIndex + 1) & "-" & ListBox1.Text
End Sub

' MECHANIZMUS PRI POSUNU SPIN BUTTONEM DOLU
Private Sub SpinButton1_SpinDown()
    If (ListBox1.ListCount > 1) Then
        If (ListBox1.ListIndex < (ListBox1.ListCount - 1)) Then
            ListBox1.ListIndex = (ListBox1.ListIndex + 1)
        End If
    End If
End Sub

' MECHANIZMUS PRI POSUNU SPIN BUTTONEM NAHORU
Private Sub SpinButton1_SpinUp()
    If (ListBox1.ListCount > 1) Then
        If (ListBox1.ListIndex >= 1) Then
            Dim oldPos As Long
            oldPos = ListBox1.ListIndex
            ListBox1.ListIndex = (ListBox1.ListIndex - 1)
        End If
    End If
End Sub
Private Sub UserForm_Activate()
    ' Pokud z nejakeho duvodu neni co vyhledavat, neprovede se nic..
    If (ThisWorkbook.searchText = "") Then
        ''MsgBox ("NENI TEXT K VYHLEDAVANI!")
        ''Me.Hide
        GoTo EndForm9
    End If
    HideBar Me
    SetTopMost Me
    Dim wb As Workbook
    Dim path As String
    path = ThisWorkbook.pathData & ThisWorkbook.FilenameRepos
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim rCount As Long
    Dim i As Long
    Dim K As Long
    '+++++++++++++++++++++++
    Dim KZM As String
    Dim PartNumber As String
    Dim Name1 As String
    Dim Name2 As String
    Dim Count As String
    Dim Repo As String
    Dim nLine As String
    
    For K = 1 To 10000
        arrKZM(K) = ""
    Next K
 
    Set wb = Workbooks.Open(path, True, True)
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    MyArr = Array(ThisWorkbook.searchText)

    With wb.Sheets(1).Range("A2:G10000")

        rCount = 0

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
                    rCount = rCount + 1
                    
                    KZM = wb.Sheets(1).Range("A" & Rng.Row).Value
                    PartNumber = wb.Sheets(1).Range("B" & Rng.Row).Value
                    Name1 = wb.Sheets(1).Range("C" & Rng.Row).Value
                    Name2 = wb.Sheets(1).Range("D" & Rng.Row).Value
                    Count = wb.Sheets(1).Range("E" & Rng.Row).Value
                    Repo = wb.Sheets(1).Range("G" & Rng.Row).Value
                    
                    If (KZM <> "") Then nLine = "KZM: " & KZM
                    If (PartNumber <> "") Then nLine = nLine & " | ID: " & PartNumber
                    If (Name1 <> "") Then nLine = nLine & " | Nazev: " & Name1
                    If (Name2 <> "") Then nLine = nLine & " " & Name2
                    If (Count <> "") Then nLine = nLine & " | Pocet: " & Count
                    If (Repo <> "") Then nLine = nLine & " | Misto: " & Repo
                    
                    arrKZM(rCount) = KZM
                    ListBox1.AddItem nLine, -1
                    Label2.Caption = rCount

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
        
    If (ListBox1.ListCount > 0) Then
        ListBox1.ListIndex = 0
        SpinButton1.Max = rCount
        SpinButton1.min = 0
        ListBox1.SetFocus
    Else
        Label2.Caption = 0
        SpinButton1.Max = 1
    End If
EndForm9:
End Sub

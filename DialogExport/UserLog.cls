VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dataPath As String
Dim wb As Workbook
Private Function InitDir() As Boolean
    On Error GoTo ErrHand
    Dim isF As Boolean
    isF = False
    If Len(Dir(dataPath, vbDirectory)) = 0 Then
       MkDir dataPath
    End If
    GoTo NormalHand
ErrHand:
    isF = False
    GoTo EndHand
NormalHand:
    isF = True
EndHand:
    InitDir = isF
End Function
Private Sub InitFile__setHeader()
Dim i As Integer
    For i = 1 To 9
    Cells(1, i).Select
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
        Select Case i
            Case 1
                Selection.ColumnWidth = 15
                Selection.Value = "DATUM"
            Case 2
                Selection.ColumnWidth = 15
                Selection.Value = "CAS"
            Case 3
                Selection.ColumnWidth = 12
                Selection.Value = "AKCE"
            Case 4
                Selection.ColumnWidth = 20
                Selection.Value = "KZM"
            Case 5
                Selection.ColumnWidth = 30
                Selection.Value = "PART NUMBER"
            Case 6
                Selection.ColumnWidth = 30
                Selection.Value = "NAZEV"
            Case 7
                Selection.ColumnWidth = 12
                Selection.Value = "POCET"
            Case 8
                Selection.ColumnWidth = 15
                Selection.Value = "MISTO"
            Case 9
                Selection.ColumnWidth = 40
                Selection.Value = "DODATKY"
        End Select
    Next i
End Sub
Private Function InitFile() As Boolean
    On Error GoTo ErrHand
    Dim isF As Boolean
    isF = False
    If Len(Dir(dataPath & ".xlsx", vbNormal)) = 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Set wb = Workbooks.Add
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "LEDEN"
        InitFile__setHeader
        wb.Sheets(1).Delete
        wb.Sheets(1).Delete
        wb.Sheets(1).Delete
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "ÚNOR"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "BØEZEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "DUBEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "KVÌTEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "ÈERVEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "ÈERVENEC"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "SRPEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "ZÁØÍ"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "ØÍJEN"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "LISTOPAD"
        InitFile__setHeader
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).name = "PROSINEC"
        InitFile__setHeader
        wb.Sheets(Month(Now)).Select
        wb.SaveAs dataPath & ".xlsx"
        wb.Close
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End If
    GoTo NormalHand
ErrHand:
    isF = False
    GoTo EndHand
NormalHand:
    isF = True
EndHand:
    InitFile = isF
End Function
'' TATO FUNKCE MUSI BYT SPUSTENA JAKO PRVNI !!!!!!!
'' ************************************************
Public Function SetClass() As Boolean
    Dim isOK As Boolean
    isOK = False
    dataPath = ThisWorkbook.pathData & ThisWorkbook.pathDataLogs
    If (InitDir = True) Then
        dataPath = ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & year(Now)
        If (InitFile = True) Then
            isOK = True
        End If
    End If
    If (isOK = False) Then
        SetClass = False
    Else
        SetClass = True
    End If
End Function
Public Function removeLog(Month As Integer, NumberLog As Integer)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Workbooks.Open ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & year(Now)
    ActiveWorkbook.Sheets(Month).Select
    Range("A" & (1 + NumberLog)).EntireRow.Delete
    ActiveWorkbook.save
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function
Public Function writeLog(Akce As Integer, KZM As String, PartNumber As String, Nazev As String, Pocet As String, Misto As String, Dodatky As String) As Boolean
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Workbooks.Open ThisWorkbook.pathData & ThisWorkbook.pathDataLogs & "\" & ThisWorkbook.GetLoginID & "_" & year(Now)
    ActiveWorkbook.Sheets(Month(Now)).Select
    Dim lastRow As Long
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    lastRow = ActiveCell.Row
    Dim i As Integer
    For i = 1 To 9
        Cells(lastRow, i).Select
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
        Selection.NumberFormat = "@"
        Select Case i
            Case 1
                Selection.NumberFormat = "dd/mm/yyyy"
                Selection.Value = Date
            Case 2
                Selection.NumberFormat = "h:mm:ss"
                Selection.Value = Time
            Case 3
                If (Akce = 0) Then
                Selection.Interior.ColorIndex = 30
                Selection.Value = "UNLOG"
                End If
                If (Akce = 1) Then
                Selection.Interior.ColorIndex = 43
                Selection.Value = "LOG"
                End If
                If (Akce = 2) Then
                Selection.Interior.ColorIndex = 42
                Selection.Value = "ODEBRAT"
                End If
                If (Akce = 3) Then
                Selection.Interior.ColorIndex = 22
                Selection.Value = "VYMAZAT"
                End If
            Case 4
                Selection.Value = KZM
            Case 5
                Selection.Value = PartNumber
            Case 6
                Selection.Value = Nazev
            Case 7
                Selection.Value = Pocet
            Case 8
                Selection.Value = Misto
            Case 9
                Selection.Value = Dodatky
        End Select
    Next i
    Cells(ActiveCell.Row + 1, 1).Select
    ActiveWorkbook.save
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
End Sub

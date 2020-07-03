VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "Sklad"
   ClientHeight    =   11295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim td As Date
Dim hour, min, sec As String
Dim saveHour, saveMin, saveSec As String

''TLACITKO ODHLASIT
Private Sub CommandButton1_Click()
    Dim writeLog As UserLog
    Set writeLog = New UserLog
    writeLog.writeLog 0, "", "", "", "", "", "" ''zapis o odhlaseni
    NoTopMost Me
    Me.Hide
End Sub
''TLACITKO HLEDAT
Private Sub CommandButton2_Click()
    UserForm8.Show
    If (ThisWorkbook.searchText <> "") Then
        UserForm9.Show
    End If
End Sub
''TLACITKO ODEBRAT (ZRYCHLENE)
Private Sub CommandButton3_Click()
    UserForm14.Show
End Sub
''TLACITKO MOJE AKTIVITA
Private Sub CommandButton4_Click()
    UserForm11.Show
End Sub
''TLACITKO REZIM SPRAVCE
Private Sub CommandButton5_Click()
NoTopMost Me
Me.Hide
UserForm15.Show
Me.Show
SetTopMostAndFullscreen Me
End Sub

Private Sub Frame1_Click()
End Sub

Private Sub UserForm_Activate()
    HideBar Me
    SetTopMostAndFullscreen Me
    Frame1.Left = ((Me.Width - Me.Left) / 2) - (Frame1.Width / 2)
    Frame1.Top = ((Me.Height - Me.Top) / 2) - (Frame1.Height / 2)
    CommandButton2.SetFocus
    ''Do horni "listy" informaci o prihlasenem uzivateli vypiseme jeho hodnost
    ''hodnost taktez rozhoduje o tom, zdali se zobrazi tlacitko SPRAVCE v menu
    ''Typicky, hodnost 0 je uzivatel, hodnost 1 je aktualne spravce///dle moznosti tohoto programu
    Select Case (ThisWorkbook.GetSU(ThisWorkbook.GetLoginID))
        Case "0"
            Label4.Caption = "TECHNIK"
        Case "1"
            Label4.Caption = "SPRÁVCE"
            CommandButton5.Visible = True
            CommandButton5.enabled = True
            Label11.Visible = True
            Label11.enabled = True
        Case Else
            Label4.Caption = "TECHNIK"
    End Select
    ''a zde vypiseme jeho prijmeni a jmeno
    Label5.Caption = ThisWorkbook.GetNameFromID(ThisWorkbook.GetLoginID)
    ''a jeho postovni ID
    Label6.Caption = ThisWorkbook.GetLoginID
End Sub


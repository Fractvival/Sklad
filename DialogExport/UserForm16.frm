VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm16 
   Caption         =   "SLOŽKA SKLADU NENÍ DOSTUPNÁ!"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9435
   OleObjectBlob   =   "UserForm16.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim sFolder As String
    sFolder = ""
    Me.Hide
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' pokud je vnitrni promenna .Show = -1, je stiknuto tlacitko OK
            sFolder = .SelectedItems(1) ' zde ulozime do promenne sFolder vybranou slozku
        End If
    End With
    If Len(sFolder) = 0 Then ' Pokud je delka textu sFolder rovna nule, ukoncime cely proces
        MsgBox "Nebyla zvolena složka pro databázi!"
        Application.Quit ' ..jelikoz se nepodarilo zvolit novou cestu, neni s cim dale pracovat
        Workbooks.Close
    Else
        If sFolder = "" Then ' ..dodatecne otestovani primo obsahu promenne sFolder, a pokud zadny neni..
            MsgBox "Nebyla zvolena složka pro databázi!"
            Application.Quit ' ..ukoncime cely proces
            Workbooks.Close
        Else
            List1.Range("B1").Value = sFolder '..do bunky B1 v sesitu List1 ulozime vyse ziskanou slozku..
            ActiveWorkbook.save '..a ulozime
            ThisWorkbook.pathData = sFolder
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
    Application.Quit
    Workbooks.Close
End Sub

Private Sub UserForm_Activate()
End Sub

Private Sub UserForm_Click()
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "Poista materiaali automaattitilaukselta"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKNappi_Click()
    Dim materialNumero As Long
    Materiaalinumero = materiaaliCombo.Value
    Sheets("Automaattitilaukset").Select
    i = 1
    Do While Materiaalinumero <> ActiveCell.Value
        Cells(i, 3).Select  'Etsitaan oikea rivi
        i = i + 1
    Loop
    i = i - 1
    Cells(i, 1).Value = ""
    Cells(i, 2).Value = ""
    Cells(i, 3).Value = ""
    Cells(i, 4).Value = ""
    Cells(i, 5).Value = ""
    Sheets("Tilaukset").Select
    Unload Me
End Sub

Private Sub peruutaNappi_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'Annetaan UserFormille alkuarvot
    materiaaliCombo.Clear
    Dim i As Integer
    For i = 0 To 1999    'valitaan toimittajien tiedoista kaikki toimittajat listaan
        With materiaaliCombo
            If Sheets("Automaattitilaukset").Cells(2 + i, 3).Value <> "" Then
            .AddItem Sheets("Automaattitilaukset").Cells(2 + i, 3).Value
            End If
        End With
    Next i
End Sub


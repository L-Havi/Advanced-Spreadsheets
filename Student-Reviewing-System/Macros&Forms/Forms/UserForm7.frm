VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "Poista k‰ytt‰j‰"
   ClientHeight    =   2350
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4230
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub peruutaButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    kayttajaTextbox.Value = ""
    nykyinenSalasanaTextbox.Value = ""
    vahvistaSalasanaTextbox.Value = ""
End Sub

Private Sub vahvistaButton_Click()

    tunnus = kayttajaTextbox.Value
    salasana = nykyinenSalasanaTextbox.Value
    vahvistaSalasana = vahvistaSalasanaTextbox.Value
    Dim oikeaTunnus As Boolean
    oikeaTunnus = False
    Dim oikeaSalasana As Boolean
    oikeaSalasana = False
    Sheets("masterdata").Visible = True
    Sheets("masterdata").Select
    oikearivi = 2
    Do While Cells(oikearivi, 3).Value <> ""
        If Cells(oikearivi, 3).Value = tunnus Then
            oikeaTunnus = True
            Exit Do
        Else
        oikearivi = oikearivi + 1
        End If
    Loop
    If Cells(oikearivi, 4).Value = salasana Then
        oikeaSalasana = True
    End If
    Sheets("Etusivu").Select
    Sheets("masterdata").Visible = False
    If oikeaTunnus = False Or oikeaSalasana = False Then
        virhe = MsgBox("Salasana tai k‰ytt‰j‰tunnus v‰‰r‰", 0, "Huomio")
    Else
        If salasana <> vahvistaSalasana Then
            virhe = MsgBox("Antamasi salasana eroaa vahvistuksesta", 0, "Virhe")
        Else
            Sheets("masterdata").Visible = True
            Sheets("masterdata").Select
            Cells(oikearivi, 1).Value = ""
            Cells(oikearivi, 2).Value = ""
            Cells(oikearivi, 3).Value = ""
            Cells(oikearivi, 4).Value = ""
            Cells(oikearivi, 5).Value = ""
            Sheets("Etusivu").Select
            Sheets("masterdata").Visible = False
            Sheets(tunnus).Visible = True
            Sheets(tunnus).Select
            oppilasrivi = 10
            Do While Cells(oppilasrivi, 13) <> ""
                oppilaanNimi = Cells(oppilasrivi, 13).Value & " " & tunnus
                Sheets(oppilaanNimi).Visible = True
                Sheets(oppilaanNimi).Delete
                oppilasrivi = oppilasrivi + 1
            Loop
            Sheets("Etusivu").Select
            Sheets(tunnus).Delete
        End If
    End If
    Unload Me
End Sub


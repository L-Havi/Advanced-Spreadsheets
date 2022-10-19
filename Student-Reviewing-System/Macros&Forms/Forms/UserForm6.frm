VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Vaihda salasana"
   ClientHeight    =   3040
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4140
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
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
    uusiSalasanaTextbox.Value = ""
    vahvistaUusiTextbox.Value = ""
End Sub

Private Sub vahvistaButton_Click()
    Dim oikeaTunnus As Boolean
    oikeaTunnus = False
    Dim oikeaSalasana As Boolean
    oikeaSalasana = False
    tunnus = kayttajaTextbox.Value
    salasana = nykyinenSalasanaTextbox.Value
    uusisalasana = uusiSalasanaTextbox.Value
    vahvistaSalasana = vahvistaUusiTextbox.Value
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
        virhe = MsgBox("Nykyinen salasana tai k‰ytt‰j‰tunnus v‰‰r‰", 0, "Huomio")
    Else
        If uusisalasana = salasana Then
            virhe = MsgBox("Uusi salasana ei voi olla sama kuin nykyinen salasana", 0, "Huomio!")
        Else
            If uusisalasana = vahvistaSalasana Then
                Sheets("masterdata").Visible = True
                Sheets("masterdata").Select
                Cells(oikearivi, 4).Value = uusisalasana
                Sheets("Etusivu").Select
                Sheets("masterdata").Visible = False
            Else
                virhe = MsgBox("Uusi salasana ja vahvistus eroavat toisistaan", 0, "Huomio!")
            End If
        End If
    End If
    Unload Me
End Sub

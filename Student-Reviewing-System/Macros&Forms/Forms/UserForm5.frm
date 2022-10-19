VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Kirjaudu Ulos"
   ClientHeight    =   1290
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2745
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub kirjauduUlosButton_Click()
        tunnus = Range("N2").Value
        oikearivi = 10
        Do While Cells(oikearivi, 13).Value <> ""
            oppilaanNimi = Cells(oikearivi, 13).Value
            oppilaanNimi = oppilaanNimi & " " & tunnus
            Sheets(oppilaanNimi).Visible = False
            oikearivi = oikearivi + 1
        Loop
        kokonaisKA = Range("R10").Value
        Sheets("masterdata").Visible = True
        Sheets("masterdata").Select
        oikearivi = 2
        Do While Cells(oikearivi, 3).Value <> ""
            If Cells(oikearivi, 3).Value = tunnus Then
                Cells(oikearivi, 5).Value = kokonaisKA
                Exit Do
            Else
                oikearivi = oikearivi + 1
            End If
        Loop
        Sheets("Etusivu").Visible = True
        Sheets("Etusivu").Select
        Sheets(tunnus).Visible = False
        Sheets("masterdata").Visible = False
        Unload Me
End Sub

Private Sub peruutaButton_Click()
    Unload Me
End Sub

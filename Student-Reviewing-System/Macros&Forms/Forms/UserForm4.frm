VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Kirjaudu sis‰‰n"
   ClientHeight    =   1940
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3315
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub peruutaButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    tunnusTextbox.Value = ""
    salasanaTextbox.Value = ""
End Sub

Private Sub vahvistaButton_Click()
    If Range("N2").Value <> "" Then
        virhe = MsgBox("Kirjaudu ensin ulos nykyiselt‰ k‰ytt‰j‰lt‰", 0, "Huomio!")
    Else
        Dim oikeaTunnus As Boolean
        oikeaTunnus = False
        Dim oikeaSalasana As Boolean
        oikeaSalasana = False
        tunnus = tunnusTextbox.Value
        salasana = salasanaTextbox.Value
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
            Sheets("masterdata").Visible = True
            Sheets("masterdata").Select
            tunnus = Cells(oikearivi, 3).Value
            Sheets(tunnus).Visible = True
            Sheets(tunnus).Select
            Sheets("Etusivu").Visible = False
            oppilasrivi = 10
            Do While Cells(oppilasrivi, 13) <> ""
                oppilaanNimi = Cells(oppilasrivi, 13).Value & " " & tunnus
                Sheets(oppilaanNimi).Visible = True
                oppilasrivi = oppilasrivi + 1
            Loop
            Unload Me
            Sheets("masterdata").Visible = False
        End If
    End If
End Sub

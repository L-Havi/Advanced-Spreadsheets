VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm10 
   Caption         =   "Muokkaa sopimusta"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "UserForm10.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub peruutaNappi_Click()

    Unload Me   'Lopetetaan lomake

End Sub
Private Sub OKNappi_Click()
If sopimusCombo.Value = "" Then
    sopimusvirhe = MsgBox("Valitse muokattava sopimusnumero", 0, "Huomio")
Else
    toimittajaNumero = Sheets("Toimittajientiedot").Range("A8:B206") 'Valitaan toimittajanumeron alue vlookupille
    sarake = 8
    sopimusNumero = sopimusCombo.Value
    toimittaja = toimittajaCombo.Value
    Materiaalinumero = materiaalinumeroTeksti.Value
    materiaaliKuvaus = materiaalikuvausTeksti.Value
    erakoko = erakokoTeksti.Value
    toimitusaika = toimitusaikaTeksti.Value
    If kappalehintaTeksti.Value = "" Then
    hinta = 0
    Else
    hinta = kappalehintaTeksti.Value
    End If
    Do While sopimusNumero <> Cells(sarake, 1).Value
        sarake = sarake + 1
    Loop
        valinta = MsgBox("Haluatko varmasti hyvaksya sopimuksen muutokset?", 1, "Muokkaa sopimusta") 'Vahvistetaan sopimuksen muokkaus
        If valinta = 1 Then
            vanhaToimittaja = Cells(sarake, 2).Value
            If vanhaToimittaja <> toimittaja Then
                If toimittaja = "" Then
                    toimittaja = vanhaToimittaja
                End If
                Sheets("Toimittajientiedot").Select
                For i = 8 To 206                    'Looppaa toimittajat lapi ja lisaa materiaalinimikkeihin 1 jos vastaa sopimuksen toimittajaan
                    If toimittaja = Cells(i, 1).Value Then
                        Cells(i, 1).Select
                        ActiveCell.Offset(0, 8).Range("A1").Select
                        ActiveCell.Value = ActiveCell.Value + 1
                        ActiveCell.Offset(0, -8).Range("A1").Select
                    End If
                    If vanhaToimittaja = Cells(i, 1).Value Then
                        Cells(i, 1).Select
                        ActiveCell.Offset(0, 8).Range("A1").Select
                        ActiveCell.Value = ActiveCell.Value - 1
                        ActiveCell.Offset(0, -8).Range("A1").Select
                    End If
                Next i
                Sheets("Sopimukset").Select
            End If
            Cells(sarake, 1).Value = sopimusNumero            'Annetaan sopimusnumerolle arvo
            If toimittaja <> "" Then
            Cells(sarake, 2).Value = toimittaja        'valitaan toimittaja
            On Error Resume Next
            toimittajanNumero = Application.WorksheetFunction.VLookup(toimittaja, toimittajaNumero, 2, False) 'Syotetaan toimittajanumero kayttaen vlookuppia toimittajasheetilla
            Cells(sarake, 3).Value = toimittajanNumero
            Else
            toimittaja = Cells(sarake, 2).Value
            toimittajanNumero = Cells(sarake, 3).Value
            End If
            If Materiaalinumero <> "" Then
            Cells(sarake, 4).Value = Materiaalinumero 'Annetaan materiaalinumero
            Else
            Materiaalinumero = Cells(sarake, 4).Value
            End If
            If materiaaliKuvaus <> "" Then
            Cells(sarake, 5).Value = materiaaliKuvaus 'Annetaan materiaalin kuvaus
            Else
            materiaaliKuvaus = Cells(sarake, 5).Value
            End If
            If erakoko <> "" Then
            Cells(sarake, 6).Value = erakoko          'Annetaan erakoko
            Else
            erakoko = Cells(sarake, 6).Value
            End If
            If toimitusaika <> "" Then
            Cells(sarake, 7).Value = toimitusaika     'Annetaan toimitusaika
            Else
            toimitusaika = Cells(sarake, 7).Value
            End If
            If myohastymissakkoCheck.Value = True Then  'Tarkistetaan onko checkbox ruksittu vai ei
                Cells(sarake, 9).Value = "Kylla"
            Else
                Cells(sarake, 9).Value = "Ei"
            End If
            If Cells(sarake, 9).Value = "Kylla" Then    'Jos myohastymissakko on kaytossa muutetaan ne
                Sheets("Myohastymissakko").Select
                sarake = sarake - 7
                Cells(sarake, 1).Value = toimittaja
                Cells(sarake, 2).Value = toimittajanNumero
                Cells(sarake, 3).Value = Materiaalinumero  'Annetaan materiaalinumerolle arvo
                Cells(sarake, 4).Value = materiaaliKuvaus
                muutos = InputBox("Anna uusi myohastymissakon maara", "Uusi myohastymissakko", 1)
                Cells(sarake, 5).Value = (muutos / 100)
                Sheets("Sopimukset").Select
            Else
                Sheets("Myohastymissakko").Select
                sarake = sarake - 7
                Cells(sarake, 1).Value = ""
                Cells(sarake, 2).Value = ""
                Cells(sarake, 3).Value = ""  'Tyhjataan poistuneet arvot
                Cells(sarake, 4).Value = ""
                Cells(sarake, 5).Value = ""
                Sheets("Sopimukset").Select
            End If
            sarake = sarake + 7
            If hinta <> 0 Then
            Cells(sarake, 10).Value = hinta 'Annetaan kappalehinta
            Else
            hinta = Cells(sarake, 10).Value
            End If
            'Lisatyn sopimuksen materiaalin lisaaminen materiaalilistaan
            Sheets("Materiaalilista").Select            'Valitaan materiaalilistan sheet
            Cells(sarake, 1).Value = sopimusNumero      'Annetaan sopimusnumerolle arvo
            Cells(sarake, 2).Value = toimittaja         'Annetaan toimittajalle arvo
            Cells(sarake, 3).Value = toimittajanNumero  'Annetaan toimittajanumerolle arvo
            Cells(sarake, 4).Value = Materiaalinumero  'Annetaan materiaalinumerolle arvo
            Cells(sarake, 5).Value = materiaaliKuvaus   'Annetaan materiaalikuvaukselle arvo
            'Lisataan skaalahintojen tiedot skaalahinnat sheetille, jos ne on annettu
            Sheets("Sopimukset").Select 'Palataan takaisin sopimukset lomakkeelle
            If kyllaOption.Value = True Then    'Valitaan onko valintalaatikko kylla/ei
                Cells(sarake, 8).Value = "Kylla"
                Sheets("Skaalahinnat").Select
                sarake = sarake - 7
                Cells(sarake, 1).Value = toimittaja
                Cells(sarake, 2).Value = toimittajanNumero
                Cells(sarake, 3).Value = Materiaalinumero
                Cells(sarake, 4).Value = materiaaliKuvaus
                Range("V40") = sarake
                Unload Me
                UserForm12.Show
            Else
                Cells(sarake, 8).Value = "Ei"
                Sheets("Skaalahinnat").Select
                sarake = sarake - 7
                Cells(sarake, 1).Value = ""
                Cells(sarake, 2).Value = ""
                Cells(sarake, 3).Value = ""
                Cells(sarake, 4).Value = ""
                Cells(sarake, 5).Value = ""
                Cells(sarake, 6).Value = ""
                Cells(sarake, 7).Value = ""
                Cells(sarake, 8).Value = ""
                Sheets("Sopimukset").Select
            End If
            Unload Me   'Lopetetaan lomake
        Else
            Unload Me   'Lopetetaan lomake
        End If
End If
End Sub

Private Sub toimittajaCombo_Change()
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    'Tyhjennetaan alkuarvo lomakkeeseen
    sopimusCombo.Clear
    toimittajaCombo.Clear
    Dim i As Integer
    For i = 9 To 1007   'valitaan sopimusten tiedoista kaikki sopimukset listaan
         With sopimusCombo
            If Sheets("Sopimukset").Cells(i, 1).Value <> "" Then
            .AddItem Sheets("Sopimukset").Cells(i, 1).Value
            End If
        End With
    Next i
    For i = 0 To 198    'valitaan toimittajien tiedoista kaikki toimittajat listaan
        With toimittajaCombo
            If Sheets("Toimittajientiedot").Cells(8 + i, 1).Value <> "" Then
            .AddItem Sheets("Toimittajientiedot").Cells(8 + i, 1).Value
            End If
        End With
    Next i
    'Tyhjennetaan alkuarvot lomakkeeseen
    materiaalinumeroTeksti.Value = ""
    materiaalikuvausTeksti.Value = ""
    erakokoTeksti.Value = ""
    toimitusaikaTeksti.Value = ""
    kappalehintaTeksti.Value = ""
    eiOption.Value = True
    myohastymissakkoCheck.Value = False
End Sub


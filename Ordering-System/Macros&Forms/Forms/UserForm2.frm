VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Lisaa sopimus"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MateriaalinumeroTeksti_Change()

End Sub

Private Sub myohastymissakkoCheck_Click()

End Sub
Private Sub peruutaNappi_Click()

    Unload Me   'Lopetetaan lomake

End Sub
Private Sub OKNappi_Click()
    Dim toimittajaNumero As Range
    Set toimittajaNumero = Sheets("Toimittajientiedot").Range("A8:B206") 'Valitaan toimittajanumeron alue vlookupille
    sopimusNumero = Range("X1").Value
    toimittaja = toimittajaCombo.Value
    On Error Resume Next
    toimittajanNumero = Application.WorksheetFunction.VLookup(toimittaja, toimittajaNumero, 2, False) 'Syotetaan toimittajanumero kayttaen vlookuppia toimittajasheetilla
    Materiaalinumero = materiaalinumeroTeksti.Value
    materiaaliKuvaus = materiaalikuvausTeksti.Value
    erakoko = erakokoTeksti.Value
    toimitusaika = toimitusaikaTeksti.Value
    hinta = kappalehintaTeksti.Value
    sarake = Range("X2").Value + 8  'Valitaan taulukosta oikea rivi
    If Cells(sarake, 1).Value <> "" Then
        'Sopimus korvaa vanhan sopimuksen
        valinta = MsgBox("Haluatko varmasti lisata uuden sopimuksen olemassaolevan paalle?", 1, "Lisaa sopimus") 'Vahvistetaan uuden sopimuksen lisaus uuden paalle
        If valinta = 1 Then
            vanhaNimike = Cells(sarake, 3).Value
            Cells(sarake, 1).Value = sopimusNumero            'Annetaan sopimusnumerolle arvo
            Cells(sarake, 2).Value = toimittaja        'valitaan toimittaja
            Cells(sarake, 3).Value = toimittajanNumero
            Cells(sarake, 4).Value = Materiaalinumero 'Annetaan materiaalinumero
            Cells(sarake, 5).Value = materiaaliKuvaus 'Annetaan materiaalin kuvaus
            Cells(sarake, 6).Value = erakoko          'Annetaan erakoko
            Cells(sarake, 7).Value = toimitusaika     'Annetaan toimitusaika
            If myohastymissakkoCheck.Value = True Then  'Tarkistetaan onko checkbox ruksittu vai ei
                Cells(sarake, 9).Value = "Kylla"
            Else
                Cells(sarake, 9).Value = "Ei"
            End If
            Cells(sarake, 10).Value = hinta 'Annetaan kappalehinta
            Range("X1").Value = Range("X1").Value + 1   'Lisataan sopimusnumeroihin yksi
            If kyllaOption.Value = True Then    'Valitaan onko valintalaatikko kylla/ei
                Cells(sarake, 8).Value = "Kylla"
                Unload Me
                UserForm3.Show
            Else
                Cells(sarake, 8).Value = "Ei"
            End If
            For i = 8 To 206                    'Looppaa toimittajat lapi ja lisaa materiaalinimikkeihin 1 jos vastaa sopimuksen toimittajaan
                Sheets("Toimittajientiedot").Select
                If toimittaja = Cells(i, 1).Value Then
                    Cells(i, 1).Select
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.Value = ActiveCell.Value + 1
                    ActiveCell.Offset(0, -8).Range("A1").Select
                End If
                If vanhaNimike = Cells(i, 2).Value Then
                    Cells(i, 1).Select
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.Value = ActiveCell.Value - 1
                    ActiveCell.Offset(0, -8).Range("A1").Select
                End If
            Next i
            'Lisatyn sopimuksen materiaalin lisaaminen materiaalilistaan
            Sheets("Sopimukset").Select                 'Palataan takaisin sopimukset lomakkeelle
            saldo = 0                                   'Asetetaan uuden materiaalin arvo nollaksi
            Sheets("Materiaalilista").Select            'Valitaan materiaalilistan sheet
            Cells(sarake, 1).Value = sopimusNumero      'Annetaan sopimusnumerolle arvo
            Cells(sarake, 2).Value = toimittaja         'Annetaan toimittajalle arvo
            Cells(sarake, 3).Value = toimittajanNumero  'Annetaan toimittajanumerolle arvo
            Cells(sarake, 4).Value = Materiaalinumero  'Annetaan materiaalinumerolle arvo
            Cells(sarake, 5).Value = materiaaliKuvaus   'Annetaan materiaalikuvaukselle arvo
            Cells(sarake, 6).Value = saldo              'Asetetaan uuden materiaalin alkusaldoksi 0
            'Lisataan skaalahintojen tiedot skaalahinnat sheetille, jos ne on annettu
            Sheets("Sopimukset").Select 'Palataan takaisin sopimukset lomakkeelle
            Unload Me   'Lopetetaan lomake
            sarake = Range("X2").Value + 8
            If Cells(sarake, 9).Value = "Kylla" Then    'Jos myohastymissakko on kaytossa avataan siihen liittyva Userform
                UserForm5.Show
            End If
        Else
            Unload Me   'Lopetetaan lomake
        End If
    Else
        'Uusi sopimus ilman, etta se menee vanhan paalle
        Cells(sarake, 1).Value = sopimusNumero            'Annetaan sopimusnumerolle arvo
        Cells(sarake, 2).Value = toimittaja        'valitaan toimittaja
        Cells(sarake, 3).Value = toimittajanNumero 'Syotetaan toimittajanumero kayttaen vlookuppia toimittajasheetilla
        Cells(sarake, 4).Value = Materiaalinumero 'Annetaan materiaalinumero
        Cells(sarake, 5).Value = materiaaliKuvaus 'Annetaan materiaalin kuvaus
        Cells(sarake, 6).Value = erakoko          'Annetaan erakoko
        Cells(sarake, 7).Value = toimitusaika     'Annetaan toimitusaika
        If myohastymissakkoCheck.Value = True Then  'Tarkistetaan onko checkbox ruksittu vai ei
            Cells(sarake, 9).Value = "Kylla"
        Else
            Cells(sarake, 9).Value = "Ei"
        End If
        Cells(sarake, 10).Value = hinta  'Annetaan kappalehinta
        Range("X1").Value = Range("X1").Value + 1   'Lisataan sopimusnumeroihin yksi
        If kyllaOption.Value = True Then    'Valitaan onko valintalaatikko kylla/ei
            Cells(sarake, 8).Value = "Kylla"
            Unload Me
            UserForm3.Show
        Else
            Cells(sarake, 8).Value = "Ei"
        End If
        sarake = Range("X2").Value + 8  'Valitaan taulukosta oikea rivi
        If Cells(sarake, 8).Value <> "Kylla" Then
            For i = 8 To 206                        'Looppaa toimittajat lapi ja lisaa materiaalinimikkeihin 1 jos vastaa sopimuksen toimittajaan
            Sheets("Toimittajientiedot").Select
                If toimittaja = Sheets("Toimittajientiedot").Cells(i, 1) Then
                    Cells(i, 1).Select
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.Value = ActiveCell.Value + 1
                    ActiveCell.Offset(0, -8).Range("A1").Select
                End If
            Next i
            'Lisatyn sopimuksen materiaalin lisaaminen materiaalilistaan
            saldo = 0                                   'Asetetaan uuden materiaalin arvo nollaksi
            Sheets("Materiaalilista").Select            'Valitaan materiaalilistan sheet
            Cells(sarake, 1).Value = sopimusNumero      'Annetaan sopimusnumerolle arvo
            Cells(sarake, 2).Value = toimittaja         'Annetaan toimittajalle arvo
            Cells(sarake, 3).Value = toimittajanNumero  'Annetaan toimittajanumerolle arvo
            Cells(sarake, 4).Value = Materiaalinumero   'Annetaan materiaalinumerolle arvo
            Cells(sarake, 5).Value = materiaaliKuvaus   'Annetaan materiaalikuvaukselle arvo
            Cells(sarake, 6).Value = saldo              'Asetetaan uuden materiaalin alkusaldoksi 0
            Sheets("Sopimukset").Select 'Palataan takaisin sopimukset lomakkeelle
            Unload Me   'Lopetetaan lomake
            sarake = Range("X2").Value + 8
            If Cells(sarake, 9).Value = "Kylla" Then    'Jos myohastymissakko on kaytossa avataan siihen liittyva Userform
                UserForm5.Show
            End If
        End If
    End If
End Sub

Private Sub toimittajaCombo_Change()
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    'Tyhjennetaan alkuarvo lomakkeeseen
    toimittajaCombo.Clear
    Dim i As Integer
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

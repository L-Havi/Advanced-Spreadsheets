Attribute VB_Name = "Module7"
Sub tallennaMateriaalitPDF()
Attribute tallennaMateriaalitPDF.VB_ProcData.VB_Invoke_Func = " \n14"
    With ActiveSheet.PageSetup
        .PrintArea = "A8:F1009"
        .Zoom = False
        .FitToPagesTall = False 'Valmistellaan taulukon alue
        .FitToPagesWide = 1
    End With
    Range("A8:F1009").ExportAsFixedFormat Type:=xlTypePDF, Filename:="materiaalit.pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub
Sub manuaalinenSaldonMuutos()
    sarake = Range("Z1").Value + 8  'Valitaan taulukosta oikea rivi
    If Cells(sarake, 4).Value = "" Then
        tarkastus = MsgBox("Valitsemallasi rivilla ei ole materiaalia", 0, "Manuaalinen saldonmuutos")
    Else
    alkuperainen = Cells(sarake, 6).Value
    muutos = InputBox("Anna uusi saldo", "Manuaalinen saldonmuutos", 1) 'Kysytaan uusi saldo
    syy = InputBox("Anna saldonmuutoksen syy", "Manuaalinen saldonmuutos", 1)   'Kysytaan muutoksen syy
    vahvistus = MsgBox("Haluatko varmasti muuttaa saldon?", 1, "Manuaalinen saldonmuutos")  'Vahvistetaan muutos
    If vahvistus = 1 Then   'Jos vahvistettu niin toteutetaan muutos
        Materiaalinumero = Cells(sarake, 4).Value
        Cells(sarake, 6).Value = muutos 'Saldo sarakkeeseen uusi saldon arvo
        Sheets("Viestit").Select    'Valitaan viestit sheet
        tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 1  'Valitaan tyhja rivi
        Cells(tyhja, 1) = Range("AB2").Value      'Annetaan paivamaara
        Cells(tyhja, 2) = Date      'Annetaan paivamaara
        Cells(tyhja, 3) = Time()    'Annetaan kellonaika
        Cells(tyhja, 4) = Materiaalinumero  'Annetaan materiaalinumero
        Cells(tyhja, 5) = syy       'Annetaan syotetty syy
        Cells(tyhja, 6) = muutos - alkuperainen 'Annetaan muutoksen maara
        Range("AB2").Value = Range("AB2").Value + 1 'Kasvatetaan viestinumeron arvoa yhdella
        Sheets("Materiaalilista").Select    'Palataan materiaalilista sheetille
    End If
    For i = 2 To 2001
        Dim tilausRaja As Integer
        skaalaArvo = 1
        Sheets("Automaattitilaukset").Select
        If Cells(i, 1).Value <> "" Then 'Tarkastetaan onko solu tyhja
            toimittaja = Cells(i, 1).Value
            toimittajanNumero = Cells(i, 2).Value
            Materiaalinumero = Cells(i, 3).Value
            materiaaliKuvaus = Cells(i, 4).Value
            tilausRaja = Cells(i, 5).Value
            Sheets("Materiaalilista").Select
            j = 8
            Do While Materiaalinumero <> ActiveCell.Value
                Cells(j, 4).Select
                j = j + 1
            Loop
            j = j - 1
            If tilausRaja > (Cells(j, 6) + Cells(j, 20)) Then 'Jos materiaali alle tilausrajan suoritetaan tilaukseen liittyva koodi
                skaalaArvo = 1
                Dim skaalat As Range
                Set skaalat = Sheets("Skaalahinnat").Range("C2:H1001")
                Sheets("Sopimukset").Select
                k = 8
                Do While Materiaalinumero <> ActiveCell.Value
                    Cells(k, 4).Select
                    k = k + 1
                Loop
                k = k - 1
                materiaaliKuvaus = Cells(k, 5).Value                  'Syotetaan materiaalikuvaus
                hinta = Cells(k, 10).Value                            'Syotetaan hinta
                sopimusNumero = Cells(k, 1).Value                    'Syotetaan sopimusnumero
                toimittaja = Cells(k, 2).Value                       'Syotetaan toimittaja
                toimittajanNumero = Cells(k, 3).Value                 'Syotetaan toimittajanumero
                skaala = Cells(k, 8).Value                           'Syotetaan skaalahinnan arvo
                tilausPaiva = Cells(k, 7).Value                      'Annetaan tilausaika
                erakoko = Cells(k, 6).Value                      'Annetaan tilausmaara
                Sheets("Tilaukset").Select
                tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 11  'Valitaan tyhja rivi
                Cells(tyhja, 1).Value = Range("Z1").Value
                Cells(tyhja, 3).Value = Date    'Annetaan arvoksi paivamaara
                Cells(tyhja, 6).Value = Materiaalinumero   'Annetaan arvoksi materiaalinumero
                Cells(tyhja, 2).Value = sopimusNumero               'Annetaan arvoksi sopimusnumero
                Cells(tyhja, 4).Value = toimittaja                  'Annetaan arvoksi toimittaja
                Cells(tyhja, 5).Value = toimittajanNumero            'Annetaan arvoksi syotetty toimittajan numero
                Cells(tyhja, 7).Value = materiaaliKuvaus            'Annetaan arvoksi syotetty materiaalikuvaus
                Cells(tyhja, 8).Value = erakoko                 'Annetaan arvoksi toimitusmaara
                If skaala = "Kylla" Then
                    On Error Resume Next
                    kymmenen = Application.WorksheetFunction.VLookup(Materiaalinumero, skaalat, 3, 0)
                    On Error Resume Next
                    viisitoista = Application.WorksheetFunction.VLookup(Materiaalinumero, skaalat, 4, 0)
                    On Error Resume Next
                    kaksikymmentaviisi = Application.WorksheetFunction.VLookup(Materiaalinumero, skaalat, 5, 0)
                    On Error Resume Next
                    kolmekymmenta = Application.WorksheetFunction.VLookup(Materiaalinumero, skaalat, 6, 0)
                    If erakoko >= kymmenen And erakoko < viisitoista Then
                        skaalaArvo = 0.9
                    ElseIf erakoko >= viisitoista And erakoko < kaksikymmentaviisi Then
                        skaalaArvo = 0.85
                    ElseIf erakoko >= kaksikymmentaviisi And erakoko < kolmekymmenta Then
                        skaalaArvo = 0.75
                    ElseIf erakoko >= kolmekymmenta Then
                        skaalaArvo = 0.7
                    End If
                End If
                Cells(tyhja, 9).Value = (hinta * erakoko * skaalaArvo)   'Annetaan arvoksi hinta
                Cells(tyhja, 10).Value = DateAdd("d", tilausPaiva, Date)   'Annetaan arvoksi syotetty toimituspaiva
                Range("Z1").Value = Range("Z1").Value + 1
                Sheets("Materiaalilista").Select
                k = 8
                Do While Materiaalinumero <> ActiveCell.Value
                    Cells(k, 4).Select
                    k = k + 1
                Loop
                k = k - 1
                Cells(k, 20).Value = Cells(k, 20).Value + erakoko
            End If
        End If
    Next i
    End If
    Sheets("Materiaalilista").Select    'Palataan materiaalilista sheetille
End Sub
Sub siirryViesteihin()
Attribute siirryViesteihin.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Viestit").Select    'Siirtyy viesteihin
End Sub
Sub poistaViesti()
    vahvistus = MsgBox("Haluatko varmasti poistaa viestin?", 1, "Viestin poisto")  'Vahvistetaan poisto
    If vahvistus = 1 Then
    sarake = Range("AB1").Value + 1  'Valitaan taulukosta oikea rivi
    Cells(sarake, 1) = ""       'Tyhjataan arvot
    Cells(sarake, 2) = ""       'Tyhjataan arvot
    Cells(sarake, 3) = ""       'Tyhjataan arvot
    Cells(sarake, 4) = ""       'Tyhjataan arvot
    Cells(sarake, 5) = ""       'Tyhjataan arvot
    Cells(sarake, 6) = ""       'Tyhjataan arvot
    Cells(sarake, 7) = ""       'Tyhjataan arvot
    Cells(sarake, 8) = ""       'Tyhjataan arvot
    End If
End Sub
Sub luoManuaalinenViesti()
    tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 1
    Do While Cells(tyhja, 1).Value <> ""
        tyhja = tyhja + 1
    Loop
    viesti = InputBox("Kirjoita viesti", "Manuaalinen viesti", 1)   'Kysytaan viesti
    Cells(tyhja, 1) = Range("AB2").Value        'Annetaan viestinumero
    Cells(tyhja, 2) = Date                      'Annetaan paivamaara
    Cells(tyhja, 3) = Time()                    'Annetaan kellonaika
    Cells(tyhja, 5) = viesti                    'Annetaan viesti
    Range("AB2").Value = Range("AB2").Value + 1 'Kasvatetaan viestinumeron arvoa yhdella
End Sub
Sub tyhjennaKaikkiViestit()
    vahvistus = MsgBox("Haluatko varmasti tyhjentaa kaikki viestit?", 1, "Viestin poisto")  'Vahvistetaan poisto
    If vahvistus = 1 Then
        Range("A2:H1000").ClearContents
        Range("AB2").Value = 1
    End If
End Sub
Sub muokkaaSkaalahintaa()
    vahvistus = MsgBox("Haluatko varmasti muuttaa skaalahintoja?", 1, "Manuaalinen skaalahinnan muutos")  'Vahvistetaan muutos
    If vahvistus = 1 Then
    syy = InputBox("Anna skaalahintojen muutoksen syy", "Manuaalinen skaalahinnan muutos", 1)   'Kysytaan muutoksen syy
    sarake = Range("Z1").Value + 1
        If Cells(sarake, 1) <> "" Then
        Materiaalinumero = Cells(sarake, 3).Value     'Kopioidaan materiaalinumero
        Sheets("Viestit").Select    'Valitaan viestit sheet
        tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 1  'Valitaan tyhja rivi
        Do While Cells(tyhja, 1).Value <> ""
            tyhja = tyhja + 1
        Loop
        Cells(tyhja, 1) = Range("AB2").Value        'Annetaan viestinumero
        Cells(tyhja, 2) = Date                      'Annetaan paivamaara
        Cells(tyhja, 3) = Time()                    'Annetaan kellonaika
        Cells(tyhja, 4) = Materiaalinumero          'Annetaan materiaalinumero
        Cells(tyhja, 5) = syy                       'Annetaan syotetty syy
        Cells(tyhja, 7) = "x"                       'Annetaan muutettu skaalahinta
        Range("AB2").Value = Range("AB2").Value + 1 'Kasvatetaan viestinumeron arvoa yhdella
        Sheets("Skaalahinnat").Select               'Palataan materiaalilista sheetille
        UserForm4.Show                              'Avataan UserForm jos kayttaja vahvistaa sen avaamisen
        Else
            virhe = MsgBox("Valitse rivi, jossa on materiaali", 0, "Huomio")
        End If
    End If
End Sub
Sub sirrySkaalahintoihin()
Attribute sirrySkaalahintoihin.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Skaalahinnat").Select 'Siirtyy skaalahinnat sheetille
End Sub
Sub muokkaaMyohastymissakkoa()
    vahvistus = MsgBox("Haluatko varmasti muuttaa myohastymissakkoa?", 1, "Manuaalinen myohastymissakon muutos")  'Vahvistetaan muutos
    If vahvistus = 1 Then
        syy = InputBox("Anna myohastymissakon muutoksen syy", "Manuaalinen myohastymissakon muutos", 1)   'Kysytaan muutoksen syy
        sarake = Range("Z1").Value + 1  'Valitaan oikea rivi
        If Cells(sarake, 1) <> "" Then
            vanha = Cells(sarake, 5).Value
            Materiaalinumero = Cells(sarake, 3).Value
            uusi = InputBox("Anna uusi myohastymissakon %-osuus hinnasta", "Manuaalinen myohastymissakon muutos", 1)  'Kysytaan uusi arvo
            Cells(sarake, 5).Value = (uusi / 100)
            Sheets("Viestit").Select    'Valitaan viestit sheet
            tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 1  'Valitaan tyhja rivi
            Do While Cells(tyhja, 1).Value <> ""
                tyhja = tyhja + 1
            Loop
            Cells(tyhja, 1) = Range("AB2").Value        'Annetaan viestinumero
            Cells(tyhja, 2) = Date                      'Annetaan paivamaara
            Cells(tyhja, 3) = Time()                    'Annetaan kellonaika
            Cells(tyhja, 4) = Materiaalinumero          'Annetaan materiaalinumero
            Cells(tyhja, 5) = syy                       'Annetaan syotetty syy
            Cells(tyhja, 8) = ((uusi / 100) - vanha)     'Annetaan myohastymissakon muutos
            Range("AB2").Value = Range("AB2").Value + 1 'Kasvatetaan viestinumeron arvoa yhdella
            Sheets("Myohastymissakko").Select    'Palataan myohastymissakon sheetille
        Else
            virhe = MsgBox("Valitse rivi, jossa on materiaali", 0, "Huomio")
        End If
    End If
End Sub

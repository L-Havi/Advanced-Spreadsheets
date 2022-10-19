VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "Tilaa materiaalia"
   ClientHeight    =   2100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Annetaan UserFormille alkuarvot
    materiaaliCombo.Clear
    Dim i As Integer
    For i = 0 To 998    'valitaan toimittajien tiedoista kaikki toimittajat listaan
        With materiaaliCombo
            If Sheets("Sopimukset").Cells(9 + i, 4).Value <> "" Then
            .AddItem Sheets("Sopimukset").Cells(9 + i, 4).Value
            End If
        End With
    Next i
    toimitusmaaraTeksti.Value = ""
End Sub
Private Sub OKNappi_Click()
    Materiaalinumero = materiaaliCombo.Value
    erakoko = toimitusmaaraTeksti.Value
    skaalaArvo = 1
    Dim skaalat As Range
    Set skaalat = Sheets("Skaalahinnat").Range("C2:H1001")
    Sheets("Sopimukset").Select
    i = 8
    Do While Materiaalinumero <> Cells(i, 4).Value
        Cells(i, 4).Select  'Etsitaan oikea rivi
        i = i + 1
    Loop
    materiaaliKuvaus = Cells(i, 5).Value                  'Syotetaan materiaalikuvaus
    hinta = Cells(i, 10).Value                            'Syotetaan hinta
    sopimusNumero = Cells(i, 1).Value                    'Syotetaan sopimusnumero
    toimittaja = Cells(i, 2).Value                       'Syotetaan toimittaja
    toimittajanNumero = Cells(i, 3).Value                 'Syotetaan toimittajanumero
    skaala = Cells(i, 8).Value                           'Syotetaan skaalahinnan arvo
    tilausPaiva = Cells(i, 7).Value                      'Annetaan tilausaika
    Sheets("Tilaukset").Select
    tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 11  'Valitaan tyhja rivi
    Cells(tyhja, 1).Value = Range("Z1").Value
    Cells(tyhja, 3).Value = Date    'Annetaan arvoksi paivamaara
    Cells(tyhja, 6).Value = Materiaalinumero   'Annetaan arvoksi materiaalinumero
    Cells(tyhja, 2).Value = sopimusNumero               'Annetaan arvoksi sopimusnumero
    Cells(tyhja, 4).Value = toimittaja                  'Annetaan arvoksi toimittaja
    Cells(tyhja, 5).Value = toimittajanNumero            'Annetaan arvoksi syotetty toimittajan numero
    Cells(tyhja, 7).Value = materiaaliKuvaus            'Annetaan arvoksi syotetty materiaalikuvaus
    Cells(tyhja, 8).Value = erakoko   'Annetaan arvoksi syotetty toimitusmaara
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
    i = 8
    Do While Materiaalinumero <> Cells(i, 4).Value
        Cells(i, 4).Select  'Etsitaan oikea rivi
        i = i + 1
    Loop
    Cells(i, 20).Value = Cells(i, 20).Value + erakoko
    Sheets("Tilaukset").Select
    'Taman jalkeen tarkistetaan tarvitaanko automaattitilauksia
    For i = 2 To 2001
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
        Do While Materiaalinumero <> Cells(j, 4).Value
            Cells(j, 4).Select  'Etsitaan oikea rivi
            j = j + 1
        Loop
        If tilausRaja > (Cells(j, 6) + Cells(j, 20)) Then 'Jos materiaali alle tilausrajan suoritetaan tilaukseen liittyva koodi
            skaalaArvo = 1
            Sheets("Sopimukset").Select
            k = 8
            Do While Materiaalinumero <> Cells(k, 4).Value
                Cells(k, 4).Select  'Etsitaan oikea rivi
                k = k + 1
            Loop
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
            If Cells(tyhja, 3) <> "" Then         'Tarkastetaan, etta rivi on oikeasti tyhja, jos ei valitaan seuraava rivi
                Do While Cells(tyhja, 3) <> ""
                    tyhja = tyhja + 1
                Loop
            End If
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
            Do While Materiaalinumero <> Cells(k, 4).Value
                Cells(k, 4).Select  'Etsitaan oikea rivi
                k = k + 1
            Loop
            Cells(k, 20).Value = Cells(k, 20).Value + erakoko
        End If
    End If
    Next i
    Sheets("Tilaukset").Select
    Unload Me
End Sub

Private Sub peruutaNappi_Click()
    Unload Me
End Sub




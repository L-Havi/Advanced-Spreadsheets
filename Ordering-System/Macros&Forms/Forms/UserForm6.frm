VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Aseta materiaali automaattitilaukselle"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Annetaan UserFormille alkuarvot
    materiaaliCombo.Clear
    Dim i As Integer
    For i = 9 To 1007    'valitaan toimittajien tiedoista kaikki toimittajat listaan
        With materiaaliCombo
            If Sheets("Sopimukset").Cells(i, 4).Value <> "" Then
                .AddItem Sheets("Sopimukset").Cells(i, 4).Value
            End If
        End With
    Next i
    alinSaldo.Value = ""
End Sub
Private Sub OKNappi_Click()
    Materiaalinumero = materiaaliCombo.Value
    Dim minimiSaldo As Integer
    minimiSaldo = alinSaldo.Value
    Sheets("Sopimukset").Select
    Dim i As Integer
    i = 8
    Do While Materiaalinumero <> Cells(i, 4).Value
        Cells(i, 4).Select  'Etsitaan oikea rivi
        i = i + 1
    Loop
    materiaaliKuvaus = Cells(i, 5).Value                 'Syotetaan materiaalikuvaus
    toimittaja = Cells(i, 2).Value                       'Syotetaan toimittaja
    toimittajanNumero = Cells(i, 3).Value                 'Syotetaan toimittajanumero
    sopimusNumero = Cells(i, 1).Value                    'Syotetaan sopimusnumero
    skaala = Cells(i, 8).Value                           'Syotetaan skaalahinnan arvo
    tilausPaiva = Cells(i, 7).Value                      'Annetaan tilausaika
    erakoko = Cells(i, 6).Value                      'Annetaan tilausmaara
    hinta = Cells(i, 10).Value                            'Syotetaan hinta
    Sheets("Automaattitilaukset").Select
    For j = 2 To 2001
        Cells(j, 3).Select
        If Cells(j, 3) = Materiaalinumero Then
            ilmoitus = MsgBox("Antamasi materiaali on jo asetettu automaattitilaukselle", 0, "Huomio")
            Exit For
        End If
    Next j
    If ilmoitus = 1 Then
        Sheets("Tilaukset").Select
        Unload Me
    Else
    tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 1  'Valitaan tyhja rivi
    If Cells(tyhja, 3) <> "" Then         'Tarkastetaan, etta rivi on oikeasti tyhja, jos ei valitaan seuraava rivi
        Do While Cells(tyhja, 3) <> ""
            tyhja = tyhja + 1
        Loop
    End If
    Cells(tyhja, 1).Value = toimittaja                  'Annetaan arvoksi toimittaja
    Cells(tyhja, 2).Value = toimittajanNumero            'Annetaan arvoksi syotetty toimittajan numero
    Cells(tyhja, 3).Value = Materiaalinumero              'Annetaan arvoksi syotetty materiaalinumero
    Cells(tyhja, 4).Value = materiaaliKuvaus            'Annetaan arvoksi syotetty materiaalikuvaus
    Cells(tyhja, 5).Value = minimiSaldo             'Annetaan arvoksi alin saldon maara milla tilataan
    Sheets("Materiaalilista").Select
    k = 8
    Do While Materiaalinumero <> Cells(k, 4).Value
        Cells(k, 4).Select  'Etsitaan oikea rivi
        k = k + 1
    Loop
    saldo = Cells(i, 6).Value
    varauma = Cells(i, 20).Value
    kokonais = saldo + varauma
    If kokonais < minimiSaldo Then
        skaalaArvo = 1
        Sheets("Tilaukset").Select
        tyhja = Application.WorksheetFunction.CountA(Range("A:A")) + 11
        skaalaArvo = 1
        Dim skaalat As Range
        Set skaalat = Sheets("Skaalahinnat").Range("C2:H1001")
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
        Cells(k, 20).Value = Cells(k, 20).Value + erakoko
        Sheets("Tilaukset").Select
        Else
            Sheets("Tilaukset").Select
        End If
    Unload Me
    End If
End Sub

Private Sub peruutaNappi_Click()
    Unload Me
End Sub


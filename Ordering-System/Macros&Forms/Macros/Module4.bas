Attribute VB_Name = "Module4"
Sub lisaaSopimus()
    UserForm2.Show 'Avaa lomakkeen uuden sopimuksen täyttöön
End Sub
Sub poistaSopimus()
    valinta = MsgBox("Haluatko varmasti poistaa sopimuksen?", 1, "Poista sopimus")  'Vahvistetaan sopimuksen poisto
    sarake = Range("X2").Value + 8  'Valitaan taulukosta oikea rivi
    If Cells(sarake, 4).Value = "" Then
        tarkastus = MsgBox("Valitsemallasi rivilla ei ole materiaalia", 0, "Poista sopimus")
    Else
    toimittaja = Cells(sarake, 2).Value      'Annetaan muuttujalle toimittajan nimi arvoksi
    Materiaalinumero = Cells(sarake, 4).Value
    skaala = Cells(sarake, 8).Value    'Annetaan muuttujalle skaalahinnan arvo
    myohastymis = Cells(sarake, 9).Value    'Annetaan muuttujalle myohastymissakon arvo
    If valinta = 1 Then 'Tarkistetaan onko valinta OK
    Cells(sarake, 1).Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select  'Tyhjataan kaikki arvot
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, -9).Range("A1").Select 'Palataan ensimmaiseen soluun
    'Poistetaan sopimuksen materiaali automaattitilauksilta
    For i = 2 To 2001                        'Looppaa toimittajat lapi ja vahenna materiaalinimikkeista 1 jos vastaa sopimuksen toimittajaan
        Sheets("Automaattitilaukset").Select
        If Materiaalinumero = Sheets("Automaattitilaukset").Cells(i, 3) Then
        Cells(i, 1).Value = ""
        Cells(i, 2).Value = ""
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
        Cells(i, 5).Value = ""
        End If
    Next i
    'Vahennetaan toimittajien materiaalinikkeista poistettu sopimus
    For i = 8 To 206                        'Looppaa toimittajat lapi ja vahenna materiaalinimikkeista 1 jos vastaa sopimuksen toimittajaan
        Sheets("Toimittajientiedot").Select
        If toimittaja = Sheets("Toimittajientiedot").Cells(i, 1) Then
            Cells(i, 1).Select
            ActiveCell.Offset(0, 8).Range("A1").Select
            ActiveCell.Value = ActiveCell.Value - 1
            ActiveCell.Offset(0, -8).Range("A1").Select
        End If
    Next i
    'Poistetaan materiaalin sopimus myos materiaalilistasta
    Sheets("Materiaalilista").Select    'Valitaan materiaalilista sheet
    For i = 9 To 1009
        Sheets("Materiaalilista").Select
        If Materiaalinumero = Sheets("Materiaalilista").Cells(i, 4) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""         'Tyhjataan kaikki arvot
            Cells(i, 5).Value = ""
            Cells(i, 6).Value = ""
        End If
    Next i
    Sheets("Sopimukset").Select         'Palataan sopimukset sheetille
    'Poistetaan materiaalin sopimus myos skaalahinnoista, jos ne oli kaytossa
    If skaala = "Kylla" Then
    Sheets("Skaalahinnat").Select
        For i = 2 To 1001
        If Materiaalinumero = Sheets("Skaalahinnat").Cells(i, 3) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""
            Cells(i, 5).Value = ""
            Cells(i, 6).Value = ""
            Cells(i, 7).Value = ""
            Cells(i, 8).Value = ""
        End If
    Next i
    End If
    Sheets("Sopimukset").Select
    'Poistetaan materiaalin sopimus myos myohastymissakoista, jos ne oli kaytossa
    If myohastymis = "Kylla" Then
        For i = 2 To 1101
        Sheets("Myohastymissakko").Select
        If Materiaalinumero = Sheets("Myohastymissakko").Cells(i, 3) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""
            Cells(i, 5).Value = ""
        End If
        Next i
        Sheets("Sopimukset").Select
    End If
    End If
    End If
End Sub


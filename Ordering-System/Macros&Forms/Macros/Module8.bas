Attribute VB_Name = "Module8"
Sub siirryAutomaattitilauksiin()
Attribute siirryAutomaattitilauksiin.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Automaattitilaukset").Select    'Siirtyy Automaattitilaukset sheetille
End Sub
Sub uusiTilaus()
    UserForm7.Show
End Sub
Sub lisaaSaapumispaiva()
    Dim myohastys As Range
    Set myohastys = Sheets("Myohastymissakko").Range("C2:E1101")
    Dim onSakko As Range
    Set onSakko = Sheets("Sopimukset").Range("D9:I1007")
    sarake = Range("Z2").Value + 11
    erakoko = Cells(sarake, 8).Value
    If Cells(sarake, 11) <> "" Then 'Varmistetaan ettei anneta samalle riville saapumispaivaa useaan kertaan
        huomio = MsgBox("Materiaalilla on jo saapumispaiva", 0, "Huomio")
    ElseIf Cells(sarake, 1) = "" Then
        huomio = MsgBox("Rivillä ei ole tilausta", 0, "Huomio")
    Else
        Cells(sarake, 11).Value = InputBox("Anna materiaalin saapumispaiva", "Saapumispaivan lisasaminen", 1)
        On Error Resume Next
        If Application.WorksheetFunction.VLookup(Cells(sarake, 6).Value, onSakko, 5, 0) Then
            If CDate(Cells(sarake, 11).Value) > CDate(Cells(sarake, 10).Value) Then 'Tarkistetaan onko myohastymissakko kaytossa ja jos on ja materiaali on myohassa annetaan sakon arvo
                On Error Resume Next
                sakko = Application.WorksheetFunction.VLookup(Cells(sarake, 6).Value, myohastys, 3, 0)
                Cells(sarake, 12).Value = (Cells(sarake, 9) * sakko)
            End If
        End If
        Materiaalinumero = Cells(sarake, 6).Value
        Sheets("Materiaalilista").Select
        i = 8
        Do While Materiaalinumero <> Cells(i, 4).Value
            Cells(i, 4).Select
            i = i + 1
        Loop
        Cells(i, 6).Value = Cells(i, 6).Value + erakoko 'Annetaan materiaalilistaan oikeat saldomaarat
        Cells(i, 20).Value = Cells(i, 20).Value - erakoko
    End If
    Sheets("Tilaukset").Select
End Sub
Sub tyhjennatilaukset()
    vahvistus = MsgBox("Haluatko varmasti poistaa tilaukset?", 1, "Tilausten tyhjennys")  'Vahvistetaan poisto
    If vahvistus = 1 Then
    Range("A12:L2011").ClearContents
    Sheets("Automaattitilaukset").Range("A2:E2001").ClearContents 'Poistetaan tiedot
    Sheets("Materiaalilista").Range("T9:T1009").ClearContents
    Range("Z1").Value = 1
    End If
End Sub
Sub asetaAutomaatti()
    UserForm6.Show
End Sub
Sub poistaAutomaatti()
    UserForm8.Show
End Sub
Sub muokkaaTilausta()
    
    UserForm9.Show
End Sub

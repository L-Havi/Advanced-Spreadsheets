Attribute VB_Name = "Module1"
Sub lisaaOppilas()
Attribute lisaaOppilas.VB_ProcData.VB_Invoke_Func = " \n14"
    If Range("N2").Value = "" Then
        virhe = MsgBox("Kirjaudu k‰ytt‰j‰lle lis‰t‰ksesi oppilaan", 0, "Huomio!")
    Else
        tunnus = Range("N2").Value
        Dim oppilaanNimi As String
        oppilaanNimi = InputBox("Anna oppilaan nimi", "Lis‰‰ oppilas", 1)
        oppilaanNimi2 = oppilaanNimi & " " & tunnus
        Sheets.Add.Name = oppilaanNimi2
        Sheets("Esimerkkitaulukko").Visible = True
        Sheets("Esimerkkitaulukko").Select
        Range("A1:AZ100").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets(oppilaanNimi2).Select
        ActiveSheet.Paste
        Range("I2").Value = oppilaanNimi
        Columns("M:N").Select
        Selection.Columns.AutoFit
        Columns("B:G").Select
        Selection.Columns.AutoFit
        Columns("G:G").Select
        Selection.ColumnWidth = 100
        Range("I2:K3").Select
        Application.CutCopyMode = False
        Range("AZ40").Value = tunnus
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Sheets(tunnus).Select
        Sheets("Esimerkkitaulukko").Visible = False
        tyhjarivi = Application.WorksheetFunction.CountA(Range("B:B")) + 10
        Do While Cells(tyhjarivi, 13).Value <> ""
            If Cells(tyhjarivi, 13).Value <> "" Then
                tyhjarivi = tyhjarivi + 1
            End If
        Loop
        Cells(tyhjarivi, 13).Value = oppilaanNimi
        Cells(tyhjarivi, 14).Value = 0
        Cells(tyhjarivi, 15).Value = 0
    End If
End Sub
Sub poistaOppilas()
    If Range("N2").Value = "" Then
        virhe = MsgBox("Kirjaudu k‰ytt‰j‰lle poistaaksesi oppilaan", 0, "Huomio!")
    Else
        tunnus = Range("N2")
        oppilaanNimi = InputBox("Anna oppilaan nimi", "Poista oppilas", 1)
        oppilaanNimi2 = oppilaanNimi & " " & tunnus
        Sheets(oppilaanNimi2).Delete
        tyhjarivi = 9
        Do While Cells(tyhjarivi, 13).Value <> oppilaanNimi
            If Cells(tyhjarivi, 13).Value <> oppilaanNimi Then
                tyhjarivi = tyhjarivi + 1
            End If
        Loop
        Cells(tyhjarivi, 13).Value = ""
        Cells(tyhjarivi, 14).Value = ""
        Cells(tyhjarivi, 15).Value = ""
        Do While Cells(tyhjarivi, 13) = "" And Cells((tyhjarivi + 1), 13) <> ""
            'Kopioidaan poistetun rivin alapuolella olevat arvot
            oppilas = Cells((tyhjarivi + 1), 13).Value
            arvioinnit = Cells((tyhjarivi + 1), 14).Value
            keskiarvo = Cells((tyhjarivi + 1), 15).Value
            'Siirret‰‰n kopioidut arvot ylemm‰s
            Cells(tyhjarivi, 13).Value = oppilas
            Cells(tyhjarivi, 14).Value = arvioinnit
            Cells(tyhjarivi, 15).Value = keskiarvo
            'Poistetaan siirretty rivi
            Cells(tyhjarivi + 1, 13).Value = ""
            Cells(tyhjarivi + 1, 14).Value = ""
            Cells(tyhjarivi + 1, 15).Value = ""
            tyhjarivi = tyhjarivi + 1
        Loop
        Cells(tyhjarivi, 13).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10649344
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells(tyhjarivi, 14).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10649344
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells(tyhjarivi, 15).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10649344
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells(tyhjarivi, 13).Select
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Cells(tyhjarivi, 14).Select
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        If Cells(10, 13).Value <> "" Then
            Cells(10, 13).Select
            TableName = ActiveCell.ListObject.Name
            Set tbl = ActiveSheet.ListObjects(TableName)
            With tbl.Range
            tbl.Resize .Resize(.CurrentRegion.rows.Count)
            End With
            'Kokonaiskeskiarvo
            i = 10
            Do While Cells(i, 14) <> ""
                kokonais = kokonais + (Cells(i, 14).Value * Cells(i, 15).Value)
                i = i + 1
            Loop
            If Range("R9").Value <> 0 Then
            Range("R10").Value = kokonais / Range("R9").Value
            Else
            Range("R10").Value = 0
            End If
        Else
        Range("R10").Value = 0
        End If
    End If
End Sub
Sub lisaaArvosana()
    UserForm1.Show
End Sub
Sub poistaArvosana()
    tunnus = Range("AZ40").Value
    poistettavaRivi = InputBox("Anna poistettavan arvioinnin numero", "Poista rivi", 1) + 2
    If Cells(poistettavaRivi, 2).Value <> "" Then
        Cells(poistettavaRivi, 2).Value = ""
        Cells(poistettavaRivi, 3).Value = ""
        Cells(poistettavaRivi, 4).Value = ""
        Cells(poistettavaRivi, 5).Value = ""
        Cells(poistettavaRivi, 6).Value = ""
        Cells(poistettavaRivi, 7).Value = ""
        'Alla looppi, jolla poistetaan tyhj‰t aukot taulukosta
        Do While Cells(poistettavaRivi, 2) = "" And Cells((poistettavaRivi + 1), 2) <> ""
            'Kopioidaan poistetun rivin alapuolella olevat arvot
            numero = Cells((poistettavaRivi + 1), 2).Value
            paivamaara = Cells((poistettavaRivi + 1), 3).Value
            kellonaika = Cells((poistettavaRivi + 1), 4).Value
            arviointiLuokka = Cells((poistettavaRivi + 1), 5).Value
            arvosana = Cells((poistettavaRivi + 1), 6).Value
            selite = Cells((poistettavaRivi + 1), 7).Value
            'Siirret‰‰n kopioidut arvot ylemm‰s
            Cells(poistettavaRivi, 2).Value = numero
            Cells(poistettavaRivi, 3).Value = paivamaara
            Cells(poistettavaRivi, 4).Value = kellonaika
            Cells(poistettavaRivi, 5).Value = arviointiLuokka
            Cells(poistettavaRivi, 6).Value = arvosana
            Cells(poistettavaRivi, 7).Value = selite
            'Poistetaan siirretty rivi
            Cells(poistettavaRivi + 1, 2).Value = ""
            Cells(poistettavaRivi + 1, 3).Value = ""
            Cells(poistettavaRivi + 1, 4).Value = ""
            Cells(poistettavaRivi + 1, 5).Value = ""
            Cells(poistettavaRivi + 1, 6).Value = ""
            Cells(poistettavaRivi + 1, 7).Value = ""
            poistettavaRivi = poistettavaRivi + 1
        Loop
        If Cells(3, 2).Value <> "" Then
           Cells(3, 2).Select
            TableName = ActiveCell.ListObject.Name
            Set tbl = ActiveSheet.ListObjects(TableName)
            With tbl.Range
            tbl.Resize .Resize(.CurrentRegion.rows.Count)
            End With
        End If
        oppilaanNimi = Range("I2").Value
        keskiarvo = Range("N14").Value
        suoritukset = Range("N13").Value
        Sheets(tunnus).Select
        tyhjarivi = 9
        Do While Cells(tyhjarivi, 13).Value <> oppilaanNimi
            If Cells(tyhjarivi, 13).Value <> oppilaanNimi Then
                tyhjarivi = tyhjarivi + 1
            End If
        Loop
        Cells(tyhjarivi, 14).Value = suoritukset
        Cells(tyhjarivi, 15).Value = keskiarvo
        'Kokonaiskeskiarvo
        i = 10
        Do While Cells(i, 14) <> ""
            kokonais = kokonais + (Cells(i, 14).Value * Cells(i, 15).Value)
            i = i + 1
        Loop
        If Range("R9").Value <> 0 Then
        Range("R10").Value = kokonais / Range("R9").Value
        Else
        Range("R10").Value = 0
        End If
        oppilaanNimi = oppilaanNimi & " " & tunnus
        Sheets(oppilaanNimi).Select
    Else
    virhe = MsgBox("Anna olemassaoleva rivi poistettavaksi", 0, "Huomio!")
    End If
End Sub
Sub muokkaaArvosanaa()
    muokattavaRivi = InputBox("Anna muokattavan rivin rivinnumero", "Muokkaa rivi‰", 1) + 2
    If Cells(muokattavaRivi, 2).Value = "" Then
        virhe = MsgBox("Anna olemassaoleva rivi muokattavaksi", 0, "Virhe")
    Else
    Range("Q1").Value = muokattavaRivi
    UserForm2.Show
    End If
End Sub
Sub siirryOppilaaseen()
    If Range("N2").Value = "" Then
        virhe = MsgBox("Kirjaudu k‰ytt‰j‰lle siirty‰ksesi oppilaaseen", 0, "Huomio!")
    Else
        tunnus = Range("N2").Value
        siirryttavaOppilas = InputBox("Anna sen oppilaan nimi, jonka arviointisivulle haluat siirty‰", "Mene sivulle", 1)
        siirryttavaOppilas = siirryttavaOppilas & " " & tunnus
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = siirryttavaOppilas Then
                Sheets(siirryttavaOppilas).Select
            End If
        Next i
    
        If ActiveSheet.Name = tunnus Then
            virhe = MsgBox("Anna olemassaoleva oppilas", 0, "Huomio!")
        End If
    End If
End Sub
Sub kirjauduUlos()
    UserForm5.Show
End Sub
Sub vaihdaSalasana()
    UserForm6.Show
End Sub

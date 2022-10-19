Attribute VB_Name = "Module5"
Sub suodataSopimus()
Attribute suodataSopimus.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A9").Select  'Valitaan solu taulukosta
    Application.CutCopyMode = False
    Range("A8:I1006").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("G3:I4"), Unique:=False   'Suodatetaan taulukko
End Sub
Sub poistaSopimussuodatus()
    Sheets("Sopimukset").Unprotect Password:="8f39mY2pSq"
    ActiveSheet.ShowAllData 'Poistetaan suodatukset
    Range("G4").Value = ""  'Nollataan arvo
    Range("H4").Value = ""  'Nollataan arvo
    Range("I4").Value = ""  'Nollataan arvo
    Sheets("Sopimukset").Protect Password:="8f39mY2pSq"
End Sub

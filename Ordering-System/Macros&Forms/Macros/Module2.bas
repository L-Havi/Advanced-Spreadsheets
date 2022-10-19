Attribute VB_Name = "Module2"
Sub tallennaNimella()
Attribute tallennaNimella.VB_ProcData.VB_Invoke_Func = " \n14"
    With ActiveSheet.PageSetup
        .PrintArea = "A7:I205"
        .Zoom = False
        .FitToPagesTall = False
        .FitToPagesWide = 1
    End With
    Range("A7:I205").ExportAsFixedFormat Type:=xlTypePDF, Filename:="toimittajat.pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub
Sub Suodatus()
Attribute Suodatus.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A8").Select  'Valitaan solu taulukosta
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("A7:I205").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("A3:D4"), Unique:=False
        'Suoritetaan erikoissuodatus
End Sub
Sub TyhjennaSuodatus()
Attribute TyhjennaSuodatus.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Toimittajientiedot").Unprotect Password:="8f39mY2pSq"
    ActiveSheet.ShowAllData 'Poistetaan suodatukset
    Range("A4").Value = ""  'Nollataan arvo
    Range("B4").Value = ""  'Nollataan arvo
    Range("C4").Value = ""  'Nollataan arvo
    Range("D4").Value = ""  'Nollataan arvo
    Sheets("Toimittajientiedot").Protect Password:="8f39mY2pSq"
End Sub

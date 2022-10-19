Attribute VB_Name = "Module6"
Sub tallennaSopimuksetPDF()
Attribute tallennaSopimuksetPDF.VB_ProcData.VB_Invoke_Func = " \n14"
    With ActiveSheet.PageSetup
        .PrintArea = "A8:J1007"
        .Zoom = False
        .FitToPagesTall = False 'Valmistellaan taulukon alue
        .FitToPagesWide = 1
    End With
    Range("A8:J1007").ExportAsFixedFormat Type:=xlTypePDF, Filename:="sopimukset.pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

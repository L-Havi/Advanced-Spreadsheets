VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Luo k‰ytt‰j‰"
   ClientHeight    =   2960
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3390
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub peruutaButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    etunimiTextbox.Value = ""
    sukunimiTextbox.Value = ""
    salasanaTextbox.Value = ""
    vahvistaTextbox.Value = ""
End Sub

Private Sub vahvistaButton_Click()
    If salasanaTextbox.Value <> vahvistaTextbox.Value Then
        virhe = MsgBox("Antamasi salasana eroaa vahvistuksesta", 0, "Virhe")
    Else
    Dim etunimiTeksti As String
    etunimiTeksti = etunimiTextbox.Value
    Dim sukunimiTeksti As String
    sukunimiTeksti = sukunimiTextbox.Value
    alku = Left(etunimiTeksti, 1)
    alku = UCase(alku)
    keski = Left(sukunimiTeksti, 2)
    keski = UCase(keski)
    Dim loppu As String
    loppu = Int(Rnd * 999)
    Dim tunnus As String
    tunnus = alku & keski & loppu
    salasana = salasanaTextbox.Value
    
    Sheets("masterdata").Visible = True
    Sheets("masterdata").Select
    tyhjarivi = Application.WorksheetFunction.CountA(Range("A:A"))
    Do While Cells(tyhjarivi, 1).Value <> ""
        If Cells(tyhjarivi, 1).Value <> "" Then
            tyhjarivi = tyhjarivi + 1
        End If
    Loop
    Cells(tyhjarivi, 1).Value = etunimiTeksti
    Cells(tyhjarivi, 2).Value = sukunimiTeksti
    Cells(tyhjarivi, 3).Value = tunnus
    Cells(tyhjarivi, 4).Value = salasana
    Cells(tyhjarivi, 5).Value = 0
    Sheets("Etusivu").Select
    Sheets("masterdata").Visible = False
    ActiveSheet.Copy After:=Worksheets(Sheets.Count)
    On Error Resume Next
    ActiveSheet.Name = tunnus
    Sheets(tunnus).Select
    Range("N2").Value = tunnus
    Sheets("Etusivu").Visible = False
    Unload Me
    viesti = MsgBox("Olet onnistuneesti luonut uuden k‰ytt‰j‰n. Tunnuksesi on " + tunnus + ". Olet nyt kirjautuneena sis‰‰n.", 0, "K‰ytt‰j‰ luotu")
    End If
End Sub

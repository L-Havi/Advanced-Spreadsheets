VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "Muokkaa toimittajaa"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7470
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Asetetaan lomakkeen alkuarvot tyhjiksi
    toimittajaCombo.Clear
    Dim i As Integer
    For i = 0 To 198    'valitaan toimittajien tiedoista kaikki toimittajat listaan
        With toimittajaCombo
            If Sheets("Toimittajientiedot").Cells(8 + i, 1).Value <> "" Then
            .AddItem Sheets("Toimittajientiedot").Cells(8 + i, 1).Value
            End If
        End With
    Next i
    PuhelinnumeroTeksti.Value = ""
    SahkopostiosoiteTeksti.Value = ""
    MaaTeksti.Value = ""
    OsoiteTeksti.Value = ""
    KaupunkiTeksti.Value = ""
    PostinumeroTeksti.Value = ""
End Sub
Private Sub OKNappi_Click()
    toimittaja = toimittajaCombo.Value
    Puhelinnumero = PuhelinnumeroTeksti.Value
    sahkopostiosoite = SahkopostiosoiteTeksti.Value
    maa = MaaTeksti.Value
    osoite = OsoiteTeksti.Value
    kaupunki = KaupunkiTeksti.Value
    postinumero = PostinumeroTeksti.Value
    If toimittaja = "" Then
        toimittajavirhe = MsgBox("Valitse muokattava toimittaja", 0, "Huomio")
    Else
        sarake = 7
        Do While toimittaja <> Cells(sarake, 1)
            sarake = sarake + 1
        Loop
        Cells(sarake, 1).Value = toimittaja           'Annetaan toimittajaan syotetty arvo
        If Puhelinnumero <> "" Then
        Cells(sarake, 3).Value = Puhelinnumero        'Annetaan puhelinnumeroon syotetty arvo
        End If
        If sahkopostiosoite <> "" Then
        Cells(sarake, 4).Value = sahkopostiosoite     'Annetaan sahkopostiin syotetty arvo
        End If
        If maa <> "" Then
        Cells(sarake, 5).Value = maa                  'Annetaan maahan syotetty arvo
        End If
        If osoite <> "" Then
        Cells(sarake, 6).Value = osoite               'Annetaan osoitteeseen syotetty arvo
        End If
        If kaupunki <> "" Then
        Cells(sarake, 7).Value = kaupunki             'Annetaan kaupunkiin syotetty arvo
        End If
        If postinumero <> "" Then
        Cells(sarake, 8).Value = postinumero          'Annetaan postinumeroon syotetty arvo
        End If
        Unload Me   'Lopetetaan lomake
    End If
End Sub

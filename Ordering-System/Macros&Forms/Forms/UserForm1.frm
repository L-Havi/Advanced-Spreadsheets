VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Lisaa toimittaja"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Asetetaan lomakkeen alkuarvot tyhjiksi
    toimittajaTeksti.Value = ""
    toimittajaNumeroTeksti.Value = ""
    PuhelinnumeroTeksti.Value = ""
    SahkopostiosoiteTeksti.Value = ""
    MaaTeksti.Value = ""
    OsoiteTeksti.Value = ""
    KaupunkiTeksti.Value = ""
    PostinumeroTeksti.Value = ""
End Sub
Private Sub OKNappi_Click()
    toimittaja = toimittajaTeksti.Value
    toimittajanNumero = toimittajaNumeroTeksti.Value
    Puhelinnumero = PuhelinnumeroTeksti.Value
    sahkopostiosoite = SahkopostiosoiteTeksti.Value 'Annetaan arvot valmiille muuttujille
    maa = MaaTeksti.Value
    osoite = OsoiteTeksti.Value
    kaupunki = KaupunkiTeksti.Value
    postinumero = PostinumeroTeksti.Value
    sarake = Range("Y1").Value + 7  'Valitaan taulukosta oikea rivi
    If Cells(sarake, 1).Value <> "" Then
        valinta = MsgBox("Haluatko varmasti lisata uuden toimittajan olemassaolevan paalle?", 1, "Lisaa toimittaja") 'Vahvistetaan uuden toimittajan lisaus uuden paalle
        If valinta = 1 Then
            Cells(sarake, 1).Value = toimittaja           'Annetaan toimittajaan syotetty arvo
            Cells(sarake, 2).Value = toimittajanNumero     'Annetaan toimittajan numeroon syotetty arvo
            Cells(sarake, 3).Value = Puhelinnumero        'Annetaan puhelinnumeroon syotetty arvo
            Cells(sarake, 4).Value = sahkopostiosoite     'Annetaan sahkopostiin syotetty arvo
            Cells(sarake, 5).Value = maa                  'Annetaan maahan syotetty arvo
            Cells(sarake, 6).Value = osoite               'Annetaan osoitteeseen syotetty arvo
            Cells(sarake, 7).Value = kaupunki             'Annetaan kaupunkiin syotetty arvo
            Cells(sarake, 8).Value = postinumero          'Annetaan postinumeroon syotetty arvo
            Cells(sarake, 9).Value = 0                                'Asettaa toimittajan materiaalinimikkeiden alkuarvoksi nolla
            Range("I2").Value = Range("I2").Value + 1                 'Lisataan toimittajien maaraa yhdella
            Unload Me   'Lopetetaan lomake
        Else
        
        Unload Me   'Lopetetaan lomake
        
        End If
    Else
        Cells(sarake, 1).Value = toimittaja           'Annetaan toimittajaan syotetty arvo
        Cells(sarake, 2).Value = toimittajanNumero     'Annetaan toimittajan numeroon syotetty arvo
        Cells(sarake, 3).Value = Puhelinnumero        'Annetaan puhelinnumeroon syotetty arvo
        Cells(sarake, 4).Value = sahkopostiosoite     'Annetaan sahkopostiin syotetty arvo
        Cells(sarake, 5).Value = maa                  'Annetaan maahan syotetty arvo
        Cells(sarake, 6).Value = osoite               'Annetaan osoitteeseen syotetty arvo
        Cells(sarake, 7).Value = kaupunki             'Annetaan kaupunkiin syotetty arvo
        Cells(sarake, 8).Value = postinumero          'Annetaan postinumeroon syotetty arvo
        Cells(sarake, 9).Value = 0                                'Asettaa toimittajan materiaalinimikkeiden alkuarvoksi nolla
        Range("I2").Value = Range("I2").Value + 1                 'Lisataan toimittajien maaraa yhdella
        Unload Me   'Lopetetaan lomake
    End If
End Sub
Private Sub peruutaNappi_Click()

    Unload Me   'Lopetetaan lomake
    
End Sub

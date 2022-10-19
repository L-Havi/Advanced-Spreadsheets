VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Anna skaalahinnat"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okButton_Click()
    sarake = Range("X2").Value + 8  'Valitaan taulukosta oikea rivi
    toimittaja = Cells(sarake, 2).Value         'Annetaan muuttujalle toimittajan arvo
    kymmenen = skaalaKymmenenTeksti.Value
    viisitoista = skaalaViisitoistaTeksti.Value
    kaksikymmentaviisi = skaalaKaksikymmentaviisiTeksti.Value
    kolmekymmenta = skaalaKolmekymmentaTeksti.Value
        For i = 8 To 206                    'Looppaa toimittajat lapi ja lisaa materiaalinimikkeihin 1 jos vastaa sopimuksen toimittajaan
            Sheets("Toimittajientiedot").Select
            If toimittaja = Cells(i, 1).Value Then
                Cells(i, 1).Select
                ActiveCell.Offset(0, 8).Range("A1").Select
                ActiveCell.Value = ActiveCell.Value + 1
                ActiveCell.Offset(0, -8).Range("A1").Select
            End If
        Next i
    'Lisatyn sopimuksen materiaalin lisaaminen materiaalilistaan
    Sheets("Sopimukset").Select                 'Palataan takaisin sopimukset lomakkeelle
    sopimusNumero = Cells(sarake, 1).Value      'Annetaan muuttujalle sopimusnumeron arvo
    toimittaja = Cells(sarake, 2).Value         'Annetaan muuttujalle toimittajan arvo
    toimittajanNumero = Cells(sarake, 3).Value  'Annetaan muuttujalle toimittajanumeron arvo
    Materiaalinumero = Cells(sarake, 4).Value  'Annetaan muuttujalle materiaalinumeronumeron arvo
    materiaaliKuvaus = Cells(sarake, 5).Value   'Annetaan muuttujalle materiaalikuvauksen arvo
    saldo = 0                                   'Asetetaan uuden materiaalin arvo nollaksi
    Sheets("Materiaalilista").Select            'Valitaan materiaalilistan sheet
    Cells(sarake, 1).Value = sopimusNumero      'Annetaan sopimusnumerolle arvo
    Cells(sarake, 2).Value = toimittaja         'Annetaan toimittajalle arvo
    Cells(sarake, 3).Value = toimittajanNumero  'Annetaan toimittajanumerolle arvo
    Cells(sarake, 4).Value = Materiaalinumero  'Annetaan materiaalinumerolle arvo
    Cells(sarake, 5).Value = materiaaliKuvaus   'Annetaan materiaalikuvaukselle arvo
    Cells(sarake, 6).Value = saldo              'Asetetaan uuden materiaalin alkusaldoksi 0
    Sheets("Sopimukset").Select 'Palataan takaisin sopimukset lomakkeelle
    sarake = Range("X2").Value + 1  'Valitaan taulukosta oikea rivi
    Sheets("Skaalahinnat").Select
    Cells(sarake, 1).Value = toimittaja
    Cells(sarake, 2).Value = toimittajanNumero
    Cells(sarake, 3).Value = Materiaalinumero
    Cells(sarake, 4).Value = materiaaliKuvaus   'Annetaan arvot skaalahinnat sheettiin
    Cells(sarake, 5).Value = kymmenen
    Cells(sarake, 6).Value = viisitoista
    Cells(sarake, 7).Value = kaksikymmentaviisi
    Cells(sarake, 8).Value = kolmekymmenta
    Sheets("Sopimukset").Select 'Palataan takaisin sopimukset lomakkeelle
    sarake = Range("X2").Value + 8
    If Cells(sarake, 9).Value = "Kylla" Then    'Jos myohastymissakko on kaytossa avataan siihen liittyva Userform
        Unload Me
        UserForm5.Show
    End If
    Unload Me   'Lopetetaan lomake
End Sub

Private Sub UserForm_Initialize()
    skaalaKymmenenTeksti.Value = ""
    skaalaViisitoistaTeksti.Value = ""
    skaalaKaksikymmentaviisiTeksti.Value = ""
    skaalaKolmekymmentaTeksti.Value = ""
End Sub

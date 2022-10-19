VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Anna myohastymissakko"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okButton_Click()
    sarake = Range("X2").Value + 8  'Valitaan taulukosta oikea rivi
    toimittaja = Cells(sarake, 2).Value
    toimittajanNumero = Cells(sarake, 3).Value
    Materiaalinumero = Cells(sarake, 4).Value 'Annetaan muuttujille arvot
    materiaaliKuvaus = Cells(sarake, 5).Value
    myohastymissakko = myohastymissakkoTeksti.Value
    sarake = Range("X2").Value + 1  'Valitaan taulukosta oikea rivi
    Sheets("Myohastymissakko").Select   'Valitaan myohastymissakon sheet
    Cells(sarake, 1).Value = toimittaja
    Cells(sarake, 2).Value = toimittajanNumero
    Cells(sarake, 3).Value = Materiaalinumero 'Annetaan arvot taulukkoon
    Cells(sarake, 4).Value = materiaaliKuvaus
    Cells(sarake, 5).Value = (myohastymissakko / 100)
    Sheets("Sopimukset").Select   'Palataan sopimuksien sheetille
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    myohastymissakkoTeksti.Value = ""
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Muokkaa skaalahintaa"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okButton_Click()
    sarake = Range("Z1").Value + 1  'Valitaan taulukosta oikea rivi
    kymmenen = skaalaKymmenenTeksti.Value
    viisitoista = skaalaViisitoistaTeksti.Value     'Annetaan muuttujille arvoja
    kaksikymmentaviisi = skaalaKaksikymmentaviisiTeksti.Value
    kolmekymmenta = skaalaKolmekymmentaTeksti.Value
    Cells(sarake, 5).Value = kymmenen
    Cells(sarake, 6).Value = viisitoista    'Sijoitetaan arvot soluihin
    Cells(sarake, 7).Value = kaksikymmentaviisi
    Cells(sarake, 8).Value = kolmekymmenta
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'Annetaan alkuarvot tekstikentille
    skaalaKymmenenTeksti.Value = ""
    skaalaViisitoistaTeksti.Value = ""
    skaalaKaksikymmentaviisiTeksti.Value = ""
    skaalaKolmekymmentaTeksti.Value = ""
End Sub


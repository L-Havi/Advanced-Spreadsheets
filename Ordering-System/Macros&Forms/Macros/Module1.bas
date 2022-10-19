Attribute VB_Name = "Module1"
Public sopimusNumero As Integer
Public toimittaja As String
Public toimittajanNumero As Long
Public Puhelinnumero As Long
Public sahkopostiosoite As String
Public osoite As String
Public kaupunki As String
Public postinumero As Integer
Public maa As String
Public Materiaalinumero As Long 'Asetetaan yleisia muuttujia, jotta datatyypit olisivat oikein makroissa
Public materiaaliKuvaus As String
Public erakoko As Long
Public toimitusaika As Long
Public skaala As String
Public myohastymissakko As String
Public saldo As Double
Public hinta As Double
Public sarake As Integer
Public skaalat As Range
Public kymmenen As Double
Public viisitoista As Double
Public kaksikymmentaviisi As Double
Public kolmekymmenta As Double

Sub toimittajaLisays()
    UserForm1.Show  'Avataan toimittajan lisayksen lomake
End Sub
Sub toimittajaPoisto()
    valinta = MsgBox("Haluatko varmasti poistaa toimittajan?", 1, "Poista toimittaja")  'Vahvistetaan toimittajan poisto
    If valinta = 1 Then 'Tarkistetaan onko valinta OK
    sarake = Range("Y1").Value + 7  'Valitaan taulukosta oikea rivi
    toimittaja = Cells(sarake, 1).Value      'Annetaan muuttujalle toimittajan nimi arvoksi
    toimittajanNumero = Cells(sarake, 2).Value
    toimittajienLukumaara = Range("I2").Value
        If toimittajienLukumaara <> 0 Then  'Tarkastetaan onko toimittajia 0, jos on ei merkita toimittajien maaraa miinusmerkkiseksi
            Range("I2").Value = toimittajienLukumaara - 1
        End If
    Cells(sarake, 1).Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select  'Tyhjataan kaikki arvot
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = ""
    ActiveCell.Offset(0, -8).Range("A1").Select 'Palataan ensimmaiseen soluun
    Sheets("Sopimukset").Select
    For i = 9 To 1007
        If Cells(i, 2).Value = toimittaja Then
            Cells(i, 1).Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select  'Tyhjataan kaikki arvot
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = ""
            ActiveCell.Offset(0, -9).Range("A1").Select 'Palataan ensimmaiseen soluun
        End If
    Next i
    'Poistetaan sopimuksen materiaali automaattitilauksilta
    For i = 2 To 2001                        'Looppaa toimittajat lapi ja vahenna materiaalinimikkeista 1 jos vastaa sopimuksen toimittajaan
        Sheets("Automaattitilaukset").Select
        If toimittajanNumero = Sheets("Automaattitilaukset").Cells(i, 2) Then
        Cells(i, 1).Value = ""
        Cells(i, 2).Value = ""
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
        Cells(i, 5).Value = ""
        End If
    Next i
    'Poistetaan materiaalin sopimus myos materiaalilistasta
    Sheets("Materiaalilista").Select    'Valitaan materiaalilista sheet
    For i = 9 To 1009
        Sheets("Materiaalilista").Select
        If toimittajanNumero = Sheets("Materiaalilista").Cells(i, 3) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""         'Tyhjataan kaikki arvot
            Cells(i, 5).Value = ""
            Cells(i, 6).Value = ""
        End If
    Next i
    'Poistetaan materiaalin sopimus myos skaalahinnoista, jos ne oli kaytossa
    Sheets("Skaalahinnat").Select
    For i = 2 To 1001
        If toimittajanNumero = Sheets("Skaalahinnat").Cells(i, 2) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""
            Cells(i, 5).Value = ""
            Cells(i, 6).Value = ""
            Cells(i, 7).Value = ""
            Cells(i, 8).Value = ""
        End If
    Next i
    'Poistetaan materiaalin sopimus myos myohastymissakoista, jos ne oli kaytossa
        For i = 2 To 1101
        Sheets("Myohastymissakko").Select
        If toimittajanNumero = Sheets("Myohastymissakko").Cells(i, 2) Then
            Cells(i, 1).Value = ""
            Cells(i, 2).Value = ""
            Cells(i, 3).Value = ""
            Cells(i, 4).Value = ""
            Cells(i, 5).Value = ""
        End If
        Next i
        Sheets("Toimittajientiedot").Select
    End If
End Sub



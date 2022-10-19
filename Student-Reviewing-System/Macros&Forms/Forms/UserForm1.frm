VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Lisää arvosana"
   ClientHeight    =   3830
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5850
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okButton_Click()
    tunnus = Range("AZ40").Value
    arviointiTyyppi = ComboBox1.Value
    arvosana = ComboBox2.Value
    If automaattiCheckBox.Value = False Then
    paivamaara = CDate(TextBox2.Value)
    kellonaika = TextBox3.Value
    Else
    paivamaara = Date
    kellonaika = Time()
    End If
    selite = TextBox1.Value
    numero = Range("P1").Value
    tyhjarivi = Application.WorksheetFunction.CountA(Range("B:B")) + 2
    Do While Cells(tyhjarivi, 2).Value <> ""
    If Cells(tyhjarivi, 2).Value <> "" Then
        tyhjarivi = tyhjarivi + 1
    End If
    Loop
    Cells(tyhjarivi, 2).Value = numero
    Range("P1").Value = Range("P1").Value + 1
    Cells(tyhjarivi, 3).Value = paivamaara
    Cells(tyhjarivi, 4).Value = kellonaika
    Cells(tyhjarivi, 5).Value = arviointiTyyppi
    Cells(tyhjarivi, 6).Value = arvosana
    Cells(tyhjarivi, 7).Value = selite
    oppilaanNimi = Range("I2").Value
    keskiarvo = Range("N14").Value
    suoritukset = Range("N13").Value
    Sheets(tunnus).Select
    tyhjarivi = 9
    Do While Cells(tyhjarivi, 13).Value <> oppilaanNimi
        If Cells(tyhjarivi, 13).Value <> oppilaanNimi Then
            tyhjarivi = tyhjarivi + 1
        End If
    Loop
    Cells(tyhjarivi, 14).Value = suoritukset
    Cells(tyhjarivi, 15).Value = keskiarvo
    'Kokonaiskeskiarvo
        i = 10
        Do While Cells(i, 14) <> ""
            kokonais = kokonais + (Cells(i, 14).Value * Cells(i, 15).Value)
            i = i + 1
        Loop
        If Range("R9").Value <> 0 Then
        Range("R10").Value = kokonais / Range("R9").Value
        Else
        Range("R10").Value = 0
        End If
    oppilaanNimi = oppilaanNimi & " " & tunnus
    Sheets(oppilaanNimi).Select
    Unload Me
End Sub

Private Sub peruutaButton_Click()
        Unload Me
End Sub

Private Sub UserForm_Initialize()
        With ComboBox1
            .AddItem "Oppitunti"
            .AddItem "Näyttö"
            .AddItem "Koe"
            .AddItem "Muu"
        End With
        With ComboBox2
            .AddItem "1"
            .AddItem "2"
            .AddItem "3"
        End With
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox1.Value = ""
    automaattiCheckBox.Value = True
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Muokkaa arvosanaa"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okButton_Click()
    tunnus = Range("AZ40").Value
    muokattavaRivi = Range("Q1").Value
    vanhaArvosana = Cells(muokattavaRivi, 6).Value
    vanhaArviointiTyyppi = Cells(muokattavaRivi, 5).Value
    arviointiTyyppi = ComboBox1.Value
    arvosana = ComboBox2.Value
    If TextBox2.Value <> "" Then
    paivamaara = CDate(TextBox2.Value)
    End If
    kellonaika = TextBox3.Value
    selite = TextBox1.Value
    If TextBox2.Value <> "" Then
    Cells(muokattavaRivi, 3).Value = paivamaara
    End If
    If TextBox3.Value <> "" Then
    Cells(muokattavaRivi, 4).Value = kellonaika
    End If
    If ComboBox1.Value <> "" Then
    Cells(muokattavaRivi, 5).Value = arviointiTyyppi
    Else
        arviointiTyyppi = Cells(muokattavaRivi, 5).Value
    End If
    If ComboBox2.Value <> "" Then
    Cells(muokattavaRivi, 6).Value = arvosana
    End If
    If TextBox1.Value <> "" Then
    Cells(muokattavaRivi, 7).Value = selite
    End If
    Range("Q1").Value = ""
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
End Sub


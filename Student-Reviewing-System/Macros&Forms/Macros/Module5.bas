Attribute VB_Name = "Module5"
Sub siirryEtusivulle()
Attribute siirryEtusivulle.VB_ProcData.VB_Invoke_Func = " \n14"
    tunnus = Range("AZ40")
    Sheets(tunnus).Select
End Sub
Sub poistaKayttaja()
    If Range("N2").Value <> "" Then
        virhe = MsgBox("Kirjaudu ulos nykyiseltä käyttäjältä poistaaksesi käyttäjän", 0, "Huomio!")
    Else
        UserForm7.Show
    End If
End Sub


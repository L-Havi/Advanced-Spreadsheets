Attribute VB_Name = "Module2"
Sub luoTunnus()
    If Range("N2").Value <> "" Then
        virhe = MsgBox("Kirjaudu ulos nykyiseltä käyttäjältä ennen uuden käyttäjän tekoa", 0, "Huomio!")
    Else
        UserForm3.Show
    End If
End Sub
Sub kirjauduSisään()
    UserForm4.Show
End Sub

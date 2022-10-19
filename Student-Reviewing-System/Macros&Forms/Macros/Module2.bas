Attribute VB_Name = "Module2"
Sub luoTunnus()
    If Range("N2").Value <> "" Then
        virhe = MsgBox("Kirjaudu ulos nykyiselt‰ k‰ytt‰j‰lt‰ ennen uuden k‰ytt‰j‰n tekoa", 0, "Huomio!")
    Else
        UserForm3.Show
    End If
End Sub
Sub kirjauduSis‰‰n()
    UserForm4.Show
End Sub

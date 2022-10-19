Attribute VB_Name = "Module3"
Sub tallennaNimellä()
Attribute tallennaNimellä.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A1:AZ100").Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Dim nimi As String
    nimi = Range("I2").Value
    Application.CutCopyMode = False
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 4")).Select
    Selection.Delete
    Sheets("Taul2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Taul3").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Taul1").Select
    Sheets("Taul1").Name = nimi
    ActiveWorkbook.SaveAs Filename:="C:\Users\Kone1\Documents\" & nimi & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End Sub


Attribute VB_Name = "Módulo1"
Sub replicar()
Dim i As Double, wb1 As Workbook, wb2 As Workbook, dirpath As String, filename As String, newname As String

Set wb1 = ActiveWorkbook
Set wb2 = Workbooks(2)

dirpath = wb2.Path & "\" & wb2.Name
wb2.Close savechanges:=False

i = 1

Do While wb1.Sheets(1).Cells(i, 1) <> ""
    Workbooks.Open dirpath
    
    Set wb2 = ActiveWorkbook
    wb2.Sheets("2.5").Cells.Find("CEP ORIGEM").Offset(1, 0) = wb1.Sheets(1).Cells(i, 2)
    
    wb2.SaveAs Replace(dirpath, "CD1", "CD" & wb1.Sheets(1).Cells(i, 1))
    
    wb2.Close savechanges:=True
    
    i = i + 1
Loop

End Sub

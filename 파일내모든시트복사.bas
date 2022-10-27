Sub CopyWorkbook()

Dim origWB As Workbook
Dim newWB As Workbook
Dim ows As Worksheet

Dim rng As Range
Dim origPath As String
Dim newPath As String
Dim myDate As String


For i = 1 To 4

origPath = "C:\Users\Jiho\OneDrive\Mercer\File.xlsx" '복사할 파일 경로
newPath = "C:\Users\Jiho\OneDrive\Mercer\NewDocs\" '저장할 폴더 경로

Set origWB = Workbooks.Open(origPath)
Set ows = origWB.Worksheets("Template")
Set rng = ows.Range("C4")
rng = ows.Range("A" & i)

Application.DisplayAlerts = False
Set newWB = ActiveWorkbook
newWB.SaveAs Filename:=newPath & rng & ".xlsx", FileFormat:=xlWorkbookDefault

Dim nws As Worksheet
For Each nws In Worksheets
With nws
.Cells.Copy
.Cells.PasteSpecial (xlPasteValues)
End With

nws.Cells.Validation.Delete

For Each nws In ActiveWorkbook.Sheets
        nws.Range("A1").ClearContents 
    Next

Next nws
newWB.Save
newWB.Close
Application.DisplayAlerts = True
Next i

End Sub

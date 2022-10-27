Sub Copy()

Dim wb As Workbook
Dim ws As Worksheet
Dim nwb As Workbook
Dim nws As Worksheet
Dim rng As Range
Dim Path As String

Set wb = ThisWorkbook
Set ws = wb.Worksheets("Position Profile") '복사할 시트 
Set rng = ws.Range("G11")
Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FldrPicker
    .Title = "저장폴더선택"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub
    myFolder = .SelectedItems(1) & "\"
                                                                
For i = 1 To 10
rng = ws.Range("A" & i) '저장할 파일명을 포함하고 있는 범위

ws.Copy

Set nwb = ActiveWorkbook
Set nws = nwb.Worksheets("Position Profile")

With nws
Cells.Copy
Cells.PasteSpecial (xlPasteValues)
End With

For Each nws In ActiveWorkbook.Sheets
        nws.Range("A1:A10").ClearContents
    Next

Application.DisplayAlerts = False
nwb.SaveAs FileName:=myFolder & rng & ".xlsx", FileFormat:=xlWorkbookDefault
nwb.Close
Application.DisplayAlerts = True

Next i

End Sub

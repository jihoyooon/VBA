Sub CombineFiles()
    Dim FilesSelected, i As Integer
    Dim tempFile As FileDialog
    Dim MainBook, sourceBook As Workbook
    Dim Sheet As Worksheet
    Dim Path As String
    Path = "C:\Users\Jiho\OneDrive\Mercer\NewDocs" '저장폴더경로
    Set MainBook = Application.ActiveWorkbook
    Set tempFile = Application.FileDialog(msoFileDialogFilePicker)
    tempFile.AllowMultiSelect = True
    FilesSelected = tempFile.Show
    For i = 1 To tempFile.SelectedItems.Count
    Workbooks.Open tempFile.SelectedItems(i)
    Set sourceBook = ActiveWorkbook
        sourceBook.Worksheets(1).Copy after:=MainBook.Sheets(MainBook.Worksheets.Count)
        ActiveSheet.Name = sourceBook.Worksheets(1).Range("K2")
        sourceBook.Close
        MainBook.Worksheets(1).Name = MainBook.Worksheets(1).Range("L2")
    Next i
    MainBook.SaveAs Filename:= Path & MainBook.Worksheets(1).Range("C6") & "_" & MainBook.Worksheets(1).Range("D6") & "_" & MainBook.Worksheets(1).Range("C7") & ".xlsx" , FileFormat:=xlWorkbookDefault 
End Sub

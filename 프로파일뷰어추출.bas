Sub 프로파일뷰어추출()

Application.ScreenUpdating = False

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nws As Worksheet
    Dim DVCell1 As Range
    Dim DVCell2 As Range
    Dim DVCell3 As Range
    Dim DVRange1 As Range
    Dim DVRange2 As Range
    Dim DVRange3 As Range
    Dim DVListItem1 As Range
    Dim DVListItem2 As Range
    Dim DVListItem3 As Range
    Dim iRow As Range
    Dim i As Long
    Dim Lastrow As Long
    
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FldrPicker
    .Title = "저장폴더선택"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub
    myFolder = .SelectedItems(1) & "\"
    End With
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("★프로파일뷰어") '복사할 시트 
    
    'DVCell1,2,3: Data Validation이 포함되어 있는 셀 
    Set DVCell1 = ws.Range("C5") 
    
    Set DVCell2 = ws.Range("F5")
    
    Set DVCell3 = ws.Range("C6")
    
    Set DVRange1 = Evaluate(DVCell1.Validation.Formula1)
    For Each DVListItem1 In DVRange1
        DVCell1.Value = DVListItem1.Value
        Set DVRange2 = Evaluate(DVCell2.Validation.Formula1)
        For Each DVListItem2 In DVRange2
            DVCell2.Value = DVListItem2.Value
            Set DVRange3 = Evaluate(DVCell3.Validation.Formula1)
            For Each DVListItem3 In DVRange3
                DVCell3.Value = DVListItem3.Value
                    ws.Copy
                    Set nwb = ActiveWorkbook
                    Set nws = nwb.Worksheets("★프로파일뷰어")
                                                                
                    With nws
                    Cells.Copy
                    Cells.PasteSpecial (xlPasteValues)
                    Columns("M:Q").ClearContents
                    Cells.Validation.Delete
                    End With

                    nws.Name = nws.Range("A1")
                    nws.Tab.ColorIndex = xlColorIndexNone
                    nws.Columns("L").Delete
                    nws.Select
                    Range("B10:J146").Select
                    With Selection
                    For lngRow = .Rows.Count To 1 Step -1
                        For Each iRow In .Rows(lngRow)
                            If WorksheetFunction.CountBlank(iRow) = iRow.Cells.Count Then
                                iRow.EntireRow.Delete
                            End If
                        Next
                    Next
                    End With
                    
                    ActiveCell = nws.Range("A1").Select
                    ActiveCell.Select
                    ActiveCell.Delete
                    
                    Application.DisplayAlerts = False
                    nwb.SaveAs Filename:=myFolder & nws.Name & ".xlsx", FileFormat:=xlWorkbookDefault
                    nwb.Close
                    Application.DisplayAlerts = True
                
            Next DVListItem3
                    
        Next DVListItem2
            
    Next DVListItem1
    
    
Application.ScreenUpdating = True

End Sub

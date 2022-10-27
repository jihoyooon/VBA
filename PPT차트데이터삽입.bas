Sub OpenPPT()

    Dim myFile As String
    Dim pptApp As PowerPoint.Application
    On Error Resume Next
    
    Set pptApp = CreateObject("Powerpoint.Application")
    
    pptApp.Visible = True
    myFile = ThisWorkbook.Sheets("PPT Macro").Range("B2").Value
    pptApp.Presentations.Open (myFile)

    a = Err.Number
    
    If Err.Number <> 0 Then GoTo mm:
    Exit Sub
    
mm:
    Msg = "해당 경로에 PPT 파일이 없습니다" & vbCr & vbCr
    Msg = Msg & "경로와 파일명을 확인바랍니다."
    MsgBox Msg, vbCritical, "실행 오류"
    
    Set pptApp = Nothing

End Sub
Sub Slide6()
    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim objSlide As PowerPoint.Slide
    Dim objChart As PowerPoint.Chart
    
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptPres = pptApp.ActivePresentation
    Set objSlide = pptPres.Slides(7) '차트가 포함되어 있는 슬라이드 번호
    
    ThisWorkbook.Sheets("앞단").Range("H5:I7").Copy '차트에 삽입할 데이터 범위 
    With objSlide.Shapes("계층").Chart.ChartData '데이터 삽입할 차트 이름 
    .Activate
    .Workbook.Worksheets("Sheet1").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    .Workbook.Close True
    End With 
    
End Sub 

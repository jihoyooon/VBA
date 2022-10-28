Sub Unprotect_worksheets()
    Dim wb As Workbook, ws As Worksheet
    Dim wPath As String, wQuan As Long, n As Long
    Dim fso As Object, folder As Object, subfolder As Object, wFile As Object
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = False
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        wPath = .SelectedItems(1)
    End With


    Set fso = CreateObject("scripting.filesystemobject")
    Set folder = fso.getfolder(wPath)
    
    wQuan = folder.Files.Count
    n = 1
    For Each wFile In folder.Files
        Application.StatusBar = "Processing folder : " & folder & ". File : " & n & " of : " & wQuan
        If Right(wFile, 4) Like "*xls*" Then
            Set wb = Workbooks.Open(wFile)
            For Each ws In wb.Sheets
                ws.Unprotect "123456" '패스워드입력
            Next
            wb.Close True
        End If
        n = n + 1
    Next
    
    For Each subfolder In folder.subfolders
        wQuan = subfolder.Files.Count
        n = 1
        For Each wFile In subfolder.Files
            Application.StatusBar = "Processing folder : " & subfolder & ". File : " & n & " of : " & wQuan
            If Right(wFile, 4) Like "*xls*" Then
                Set wb = Workbooks.Open(wFile)
                For Each ws In wb.Sheets
                ws.Unprotect "123456" '패스워드입력
                Next
                wb.Close True
            End If
            n = n + 1
        Next
    Next
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Set fso = Nothing: Set folder = Nothing: Set wb = Nothing
    
    MsgBox "End"
End Sub

Sub Decrypt()

Dim FileNames As Variant, i As Integer, j As Integer
Dim TWB As Workbook, aWB As Workbook
Dim wPath As String
Dim fso As Object, folder As Object
Dim pw As String

Set TWB = ThisWorkbook
pw = TWB.Worksheets(1).Range("B1").Value
Application.ScreenUpdating = False

MsgBox ("파일을 선택하세요")

FileNames = Application.GetOpenFilename(FileFilter:="Excel Filter (*.xlsx), *xlsx", Title:="Open File(s)", MultiSelect:=True)

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub
    wPath = .SelectedItems(1)
End With

Set fso = CreateObject("scripting.filesystemobject")
Set folder = fso.getfolder(wPath)
    

For i = 1 To UBound(FileNames)

Workbooks.Open FileNames(i), Password:=pw
ActiveWorkbook.SaveAs Filename:=folder & "\" & ActiveWorkbook.Name, Password:=""
ActiveWorkbook.Close


Next i
Application.ScreenUpdating = True
Application.StatusBar = False

MsgBox "완료"

End Sub

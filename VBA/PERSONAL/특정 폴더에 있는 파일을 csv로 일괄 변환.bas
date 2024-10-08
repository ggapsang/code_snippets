Sub XLStoCSV()
    
    Application.DisplayAlerts = False
    Application.StatusBar = True
    Application.ScreenUpdating = False
    
    Dim xFd As FileDialog
    Dim xSPath As String
    Dim xExcelFile As String
    Dim xWsheet As String
    Application.DisplayAlerts = False
    Application.StatusBar = True
    xWsheet = ActiveWorkbook.Name
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    xFd.Title = "Select a folder:"
    If xFd.Show = -1 Then
        xSPath = xFd.selectedItems(1)
    Else
        Exit Sub
    End If
    If Right(xSPath, 1) <> "\" Then xSPath = xSPath + "\"
    xExcelFile = Dir(xSPath & "*.xls*")
    Do While xExcelFile <> ""
        Application.StatusBar = "Converting: " & xExcelFile
        Workbooks.Open fileName:=xSPath & xExcelFile
        ' ".xls" 및 ".xlsx" 확장자 모두 ".csv"로 변경
        Dim newFileName As String
        'newFileName = Replace(xSPath & xExcelFile, ".xls", ".csv", vbTextCompare)
        newFileName = Replace(xSPath & xExcelFile, ".xlsx", ".csv", vbTextCompare)
        
        ActiveWorkbook.SaveAs newFileName, xlCSV
        ActiveWorkbook.Close
        Windows(xWsheet).Activate
        xExcelFile = Dir
    Loop
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub



    Dim fd As FileDialog
    Dim folderPath As String
    Dim filePath As String
    Dim fileList() As String
    Dim i As Integer

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "폴더 선택"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) '폴더 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With
    Debug.Print folderPath
    
    
    filePath = Dir(folderPath & "\*.xlsx")
    i = 0
    Do While filePath <> ""
        ReDim Preserve fileList(i)
        fileList(i) = filePath
        filePath = Dir
        i = i + 1
    Loop
    Debug.Print UBound(fileList)

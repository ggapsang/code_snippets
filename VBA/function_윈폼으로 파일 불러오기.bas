Function GetWorkbookFromUser() As Workbook
    Dim wb As Workbook
    Dim filePath As String
    Dim fileName As String
    Dim isWorkbookOpen As Boolean
        
    ' 파일 선택 창 표시
    filePath = Application.GetOpenFilename("엑셀 파일(*.xls;*.xlsx;*.xlsb;*.xlsm), *.xls;*.xlsx;*.xlsb;*.xlsm", , "파일 선택", , False)
    
    ' 파일 선택 창에서 취소 버튼을 누른 경우
    If filePath = "False" Then
        MsgBox "취소", vbExclamation
        Set GetWorkbookFromUser = Nothing
        Exit Function
    End If
    
    ' 선택한 파일의 이름 가져오기
    fileName = Dir(filePath)
    
    ' 파일이 이미 열려 있는지 확인
    isWorkbookOpen = False
    For Each wb In Workbooks
        If wb.Name = fileName Then
            isWorkbookOpen = True
            Set GetWorkbookFromUser = wb
            Exit Function
        End If
    Next wb
    
    ' 파일이 열려 있지 않으면 열기
    If Not isWorkbookOpen Then
        Set wb = Workbooks.Open(filePath)
        Set GetWorkbookFromUser = wb
    End If

End Function

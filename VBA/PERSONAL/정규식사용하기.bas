Sub UseRR(pattern As String, replacement As String)
    Dim ws As Worksheet
    Dim cell As Range
    Dim regex As Object
    Dim inputText As String
    Dim selectedRange As Range
    
    ' 정규식 객체 생성
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.pattern = pattern
    regex.Global = True
    
    ' 작업 범위 선택
    Set selectedRange = Application.InputBox("작업 범위 선택:", "Range Selector", Type:=8)
    
    ' 필터링된 셀만 대상으로 작업
    For Each cell In selectedRange.SpecialCells(xlCellTypeVisible)
        If Not IsEmpty(cell.value) Then
            inputText = cell.value
            ' 패턴이 매칭되면 교체
            If regex.test(inputText) Then
                cell.value = regex.Replace(inputText, replacement)
            End If
        End If
    Next cell
    
    MsgBox "변경이 완료되었습니다!"
End Sub

Sub UseRe()
    UserForm1.Show
End Sub

' 매크로 개요:
' - 특정 시트에서 특정 셀에 값을 입력하면, 특정 시트에 표 형태로 자동으로 데이터를 추가합니다.
' - 특정 시트의 표는 Excel의 "테이블" 기능을 사용해 관리되며, 표의 동적 크기 조정 기능을 활용합니다.

Sub ExportDB()
    Dim wsInput As Worksheet
    Dim wsDB As Worksheet
    Dim dbTable As ListObject
    Dim nextRow As ListRow
    Dim isEmptyTable As Boolean
    
    Set wsSource = ThisWorkbook.Sheets("입력시트")
    Set wsDB = ThisWorkbook.Sheets("데이터베이스_표_계약관리")

    '테이블 참조 설정
    On Error Resume Next
    Set dbTable = wsDB.ListObjects("계약관리_DB") '
    On Error GoTo 0

    ' 테이블이 없으면 종료
    If dbTable Is Nothing Then
        MsgBox "'데이터베이스_표_계약관리 시트'에 '계약관리_DB' 테이블이 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 비어있는 테이블인지 확인
    isEmptyTable = (dbTable.ListRows.Count = 0)
    
    ' 새로운 행 추가
    If isEmptyTable Then
        Set nextRow = dbTable.ListRows.Add(1) ' 첫 번째 행에 추가
    Else
        Set nextRow = dbTable.ListRows.Add ' 마지막 행 다음에 추가
    End If

    ' 특정 열(예: "계약번호")에 값 입력
    nextRow.Range(1, dbTable.ListColumns("계약번호").Index).Value = wsSource.Range("A1").Value ' 입력 시트의 A1 값을 복사

End Sub

' 일반 시트 셀과 테이블의 주요 차이점:
' 1. 테이블은 동적 범위 관리가 가능하여 새로운 행이나 열이 추가되면 자동으로 범위에 포함됩니다.
' 2. 테이블 내 셀의 서식, 계산 및 이름 참조가 편리하게 관리됩니다.
' 3. 테이블은 이름 기반으로 관리되며, VBA 코드에서 ListObject로 참조됩니다.

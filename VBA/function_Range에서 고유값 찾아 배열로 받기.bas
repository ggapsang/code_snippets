Function GetUniqueValues(rng As Range) As Variant

    Dim cell As Range
    Dim uniqueDict As Object
    Dim uniqueArray() As Variant
    Dim key As Variant
    Dim i As Integer

    'Dictionary 객체 생성
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    '범위의 각 셀에 대해 고유값을 Dictionary에 추가
    For Each cell In rng
        If Not IsEmpty(cell.Value) And Not uniqueDict.Exists(cell.Value) Then
            uniqueDict.Add cell.Value, Nothing
        End If
    Next cell

    'Dictionary의 키를 배열로 변환
    i = 0
    ReDim uniqueArray(0 To uniqueDict.Count - 1)
    For Each key In uniqueDict.Keys
        uniqueArray(i) = key
        i = i + 1
    Next key

    '결과 배열 반환
    GetUniqueValues = uniqueArray

End Function

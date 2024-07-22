Function INSTRVLOOKUP(lookupValue As String, rngSearch As Range, rngReturn As Range) As Variant

'' 찾을 값이 있는 범위를 배열에 저장
    
    Dim arrSearch As Variant
    arrSearch = rngSearch.value

'' 반환할 값이 있는 범위를 배열에 저장

    Dim arrReturn As Variant
    arrReturn = rngReturn.value

'' 찾을 범위를 순회하면서, lookupValue를 포함하는 문자열이 배열에 있는지 확인하고, 포함하는 문자열이 있을 경우, 반환할 값이 있는 배열의 해당 인덱스를 저장

    For i = LBound(arrSearch, 1) To UBound(arrSearch, 1)
        If Instr(arrSearch(i,1), lookupValue) > 0 Then
            INSTRVLOOKUP = arrReturn(i, 1)
            Exit Function
        End If
    Next i

    INSTRVLOOKUP = CVERrr(xlErrNA)

End Function

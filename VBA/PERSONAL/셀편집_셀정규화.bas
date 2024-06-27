Sub BreakandFill()


    Dim header_row As Integer
    Dim cardNo_col As Integer
    Dim cardNo_col_chr As String
    Dim cell As Range
    
    Dim replaceWith As String
    
    replaceWith = InputBox("구분자 설정 : ", "Input Required", ", ")
    
    
    Application.DisplayAlerts = False

    For Each cell In Selection.Cells
        Dim value As Variant
        value = cell.value
        If InStr(1, value, vbLf) > 0 Then
            value = Replace(value, vbLf, replaceWith)
            cell.value = value
        ElseIf InStr(1, value, vbCr) > 0 Then
            value = Replace(value, vbCr, replaceWith)
            cell.value = value
        End If
    Next cell

    For Each cell In Selection.Cells
        If cell.MergeCells Then
            Dim mergeRange As Range
            Set mergeRange = cell.MergeArea
            cell.UnMerge
            value = cell.value
            value = Replace(value, vbCrLf, replaceWith)
            mergeRange.value = value
        End If
    Next cell
    
End Sub




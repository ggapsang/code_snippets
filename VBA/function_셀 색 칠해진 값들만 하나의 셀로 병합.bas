Function COLORCONNECT(rng As Range, clr As Range) As String
    Dim cell As Range
    Dim result() As String
    Dim count As Integer
    
    ' Initialize count
    count = 0
    
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Interior.ColorIndex = clr.Interior.ColorIndex Then
            ' Resize the result array and add the cell value
            ReDim Preserve result(count)
            result(count) = cell.Value
            count = count + 1
        End If
    Next cell
    
    ' Join the array elements with " & " as the delimiter
    If count > 0 Then
        COLORCONNECT = Join(result, " & ")
    Else
        COLORCONNECT = ""
    End If
End Function

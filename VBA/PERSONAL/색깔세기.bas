Function ColorCount(rng As Range, colorCell As Range) As Integer
    Dim cell As Range
    Dim count As Integer
    Dim targetColorIndex As Long
    
    ' Get the color index from the specified colorCell
    targetColorIndex = colorCell.Interior.ColorIndex
    
    ' Initialize count
    count = 0
    
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Interior.ColorIndex = targetColorIndex Then
            ' Increment count for each cell matching the target color
            count = count + 1
        End If
    Next cell
    
    ' Return the count of cells with matching color
    ColorCount = count
End Function


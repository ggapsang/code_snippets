Function FindColLetter(hdr_row As Integer, search_value As Variant) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String
    
    Set search_rng = ActiveSheet.Rows(hdr_row)
    
    Set found_cell = search_rng.Find(What:=search_value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found_cell Is Nothing Then
        col_letter = Replace(found_cell.Cells.Address(False, False), hdr_row & "", "")
        FindColLetter = col_letter
    Else
        FindColLetter = "Value not found."
    End If

End Function

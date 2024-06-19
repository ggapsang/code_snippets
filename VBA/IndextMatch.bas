Sub indexmatch()

    Dim index_col As String
    Dim match_value_col As String
    Dim match_range_col As String
    Dim rw As Integer
    Dim hdr_row As Integer
    Dim last_row As Integer


    Dim i As Integer
    Dim match_value As Variant
    
    index_col = "AY"
    match_value_col = "BM"
    match_range_col = "AX"
    rw = 12
    hdr_row = 11
    last_row = 2882

    match_value = Range(match_value_col & rw).Value

    For i = hdr_row + 1 To last_row
        If match_value = Range(match_range_col & i).Value Then
           Range("BH" & rw).Value = Range(index_col & i).Value
        Else
        End If
    Next i
        
    
End Sub

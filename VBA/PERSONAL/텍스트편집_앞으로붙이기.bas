Sub GaEun2()


    Dim header_row As Integer
    Dim cardNo_col As Integer
    Dim cardNo_col_chr As String
    Dim cell As Range
    
    Dim replaceWith As String
    
    replaceWith = InputBox("앞으로 붙이기 :")
    
    
    Application.DisplayAlerts = False

    For Each cell In Selection.Cells
        Dim value As Variant
        value = cell.value
        cell.value = replaceWith & value
    Next cell

End Sub



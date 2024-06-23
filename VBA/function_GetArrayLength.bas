Function GetArrayLength(arr As Variant) As Integer
    GetArrayLength = UBound(arr) - LBound(arr) + 1
End Function

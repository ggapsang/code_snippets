Sub DictionaryWithNestedLoop()

    ' Create a new Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Create a first Collection and add items
    Dim coll1 As Collection
    Set coll1 = New Collection
    coll1.Add "Item1"
    coll1.Add "Item2"
    coll1.Add "Item3"
    
    ' Add the first Collection to the Dictionary with key "Key1"
    dict.Add "Key1", coll1
    
    ' Create a second Collection and add items
    Dim coll2 As Collection
    Set coll2 = New Collection
    coll2.Add "ItemA"
    coll2.Add "ItemB"
    
    ' Add the second Collection to the Dictionary with key "Key2"
    dict.Add "Key2", coll2
    
    ' Loop through the Dictionary using index
    Dim i As Integer
    Dim j As Integer
    Dim dictKeys As Variant
    dictKeys = dict.Keys
    
    For i = 0 To dict.Count - 1
        ' Print the Dictionary key
        Debug.Print "Dictionary Key: " & dictKeys(i)
        
        ' Loop through the Collection associated with the current key
        For j = 1 To dict(dictKeys(i)).Count
            ' Print the Collection item
            Debug.Print "    Collection Item: " & dict(dictKeys(i)).Item(j)
        Next j
    Next i

End Sub


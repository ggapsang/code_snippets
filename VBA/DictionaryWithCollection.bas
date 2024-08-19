
Sub DictionaryWithCollection()

    ' Create a new Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Create a new Collection
    Dim coll As Collection
    Set coll = New Collection
    
    ' Add items to the Collection
    coll.Add "Item1"
    coll.Add "Item2"
    coll.Add "Item3"
    
    ' Add the Collection as a value to the Dictionary
    dict.Add "Key1", coll
    
    ' You can add another collection to a different key
    Dim coll2 As Collection
    Set coll2 = New Collection
    coll2.Add "ItemA"
    coll2.Add "ItemB"
    
    dict.Add "Key2", coll2
    
    ' Example of accessing items in the collection stored in the Dictionary
    Dim i As Integer
    For i = 1 To dict("Key1").Count
        Debug.Print dict("Key1").Item(i)
    Next i
    
    ' Output another collection's items
    For i = 1 To dict("Key2").Count
        Debug.Print dict("Key2").Item(i)
    Next i

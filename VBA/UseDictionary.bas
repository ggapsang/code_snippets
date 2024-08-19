Sub UseDictionary()
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    ' Adding items to the dictionary
    dict.Add "Key1", "Value1"
    dict.Add "Key2", "Value2"
    dict.Add "Key3", "Value3"

    ' Accessing an item by its key
    MsgBox dict("Key1")  ' Outputs: Value1

    ' Checking if a key exists
    If dict.Exists("Key2") Then
        MsgBox "Key2 exists in the dictionary."
    End If

    ' Removing an item
    dict.Remove "Key3"

    ' Looping through the dictionary
    Dim key As Variant
    For Each key In dict.Keys
        MsgBox "Key: " & key & ", Value: " & dict(key)
    Next key

    ' Clear the dictionary
    dict.RemoveAll

    ' Destroy the dictionary object
    Set dict = Nothing
End Sub

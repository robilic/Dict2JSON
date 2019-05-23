'
' Convert a Dictionary type to JSON
' Supports nested dictionaries
'
Function Dict2JSON(dict As Dictionary)
    
    Dim str As String
    str = " { "
    For Each key In dict.keys
        ' write this key
        ' "somekey" :
        str = str & Chr(34) & key & Chr(34) & ": "
        
        If IsArray(dict(key)) Then
            ' arrays are wrapped in square brackets
            str = str & " [ "
            
            Dim element As Integer
            For element = 1 To UBound(dict(key))
                ' this is the only way it seems to work
                ' dict(key)(x) is not recognized as an dict on it's own
                Dim d As Object
                Set d = dict(key)(element)
                str = str & Dict2JSON(d)
                str = str & ", "
            Next
            
            str = Left(str, Len(str) - 2)
            str = str & " ] "
        ElseIf TypeOf dict(key) Is Dictionary Then
            ' nested dictionary requires recursion
            str = str & Dict2JSON(dict(key))
        Else
            ' just a regular value, write it out:
            If VarType(dict(key)) = vbString Then
                ' enclose strings in quotes
                str = str & Chr(34) & dict(key) & Chr(34)
            Else
                str = str & dict(key)
            End If
        End If
        
        str = str & ", "
    Next
    ' remove the last comma+space
    str = Left(str, Len(str) - 2)
    str = str & " } "
    Dict2JSON = str

End Function

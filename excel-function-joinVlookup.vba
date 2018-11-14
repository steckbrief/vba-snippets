' This function allows to lookup a value in a range and join it by jointext
Public Function joinVlookup(lookupval, lookuprange As Range, indexcol As Long, jointext As String)
    Dim x As Range
    Dim result As String
    result = ""
    ' Support for new line as join text
    If "\n" = jointext Then
        jointext = Chr(10)
    End If
    For Each x In lookuprange
        If x = lookupval Then
            nextval = x.Offset(0, indexcol - 1)
            If "" = result Then
                result = nextval
            Else
                result = result & jointext & nextval
            End If
        End If
    Next x
    joinVlookup = result
End Function

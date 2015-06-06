' NOT FINISHED
Public Class Sort

    Shared Function ReverseArray(ByVal uArray() As String)
        Array.Reverse(uArray)
        Return uArray(uArray.Length - 1)
    End Function

    Shared Function SortArrayAlphabetically(ByVal uArray() As String)
        Array.Sort(uArray)
        Return uArray(uArray.Length - 1)
    End Function
End Class
' NOT FINISHED
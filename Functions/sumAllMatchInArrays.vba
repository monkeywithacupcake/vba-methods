Function sumAllMatchInArrays(inArrs As Variant, lookupCol As Integer, returnCol As Integer, match As String)
    'Function is intended to get an overall sum for a thing, 
    'for example, if there are many tabs that have similar data, and you want the total for a particular match across all tabs
    'first read the tab data into arrays and then pass an array of those arrays to this function
    'inArrs is a one dimensional array of arrays
    'lookupCol and returnCol are identified by the calling Sub or Function
    'lookupCol is the Column number where we are looking to find match
    'returnCol is the Column number on which we are summing
    Dim i As Long, a As Long, arr As Variant, oVal As Double
    oVal = 0
    For a = LBound(inArrs) To UBound(inArrs)
        arr = inArrs(a)
        For i = LBound(arr, 1) To UBound(arr, 1)
            If CStr(arr(i, lookupCol) = match) Then
                oVal = oVal + arr(i, returnCol)
            End If
        Next i
    Next a
    sumAllMatch = oVal
End Function

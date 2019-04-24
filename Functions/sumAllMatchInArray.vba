Function sumAllMatchInArray(inArr As Variant, lookupCol As Integer, returnCol As Integer, match As String)
    'Function is intended to get an overall sum for a thing from an array
    'inArr is a two dimensional array
    'lookupCol and returnCol are identified by the calling Sub or Function
    'lookupCol is the Column number where we are looking to find match
    'returnCol is the Column number on which we are summing
    Dim i As Long, oVal As Double
    oVal = 0
    For i = LBound(inArr, 1) To UBound(inArr, 1) ' for each row
        If CStr(inArr(i, lookupCol) = match) Then
            oVal = oVal + inArr(i, returnCol)
        End If
    Next i
    sumAllMatchInArray = oVal
End Function

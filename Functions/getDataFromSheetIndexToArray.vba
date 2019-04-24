Function getDataFromSheetIndexToArray(wbkPath As String, wksIndex As Integer)
    'wbkPath is a string of the Path to the file. like "C://user/folder/blah.xlsx"
    'wksIndex is expected to be an index (like 1) or 0 if you want to get the most recent tab
    Dim wbk As Workbook, wks As Worksheet, lastrow As Long, lastcol As Integer, ttarr As Variant
    Set wbk = Workbooks.Open(Filename:=wbkPath, ReadOnly:=True)
    If wksIndex <> 0 Then
        Set wks = wbk.Sheets(wksIndex)
    Else ' figure out the last worksheet
        Set wks = wbk.Sheets(wbk.Sheets.Count)
    End If
    With wks
        lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        lastcol = .Cells(2, .Columns.Count).End(xlToLeft).Column
        ttarr = .Range("A2").Resize(lastrow, lastcol).Value2
    End With
    wbk.Close
    getDataFromSheetIndexToArray = ttarr
End Function

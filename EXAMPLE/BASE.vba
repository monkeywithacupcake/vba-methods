
Option Explicit     ' all vars must be dimmed
Option Base 1       ' arrays start with 1
Option Compare Text ' That is, "AAA" is equal to "aaa".
'
'' BASE MODULE
'
' BASE modules contain functions and subs that are not file specific.
' I recommend putting all reusable functions in a BASE Module that you
' can copy and paste into new projects

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' SpeedUp Functions
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub startProcessing()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub
Public Sub endProcessing()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operations
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub pauseRunning(ptime As Long)
    Dim waitTime
    waitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + ptime)
    Application.Wait waitTime
End Sub

Public Sub updateLRDate(wbk As Workbook, fudname As String)
    wbk.Names(fudname).RefersToRange.Value = Now()
End Sub

Public Sub sortWbTabsDescending(wb As Workbook)
 Dim i As Long, j As Long
For i = 1 To wb.Sheets.Count
    For j = 1 To wb.Sheets.Count - 1
        If UCase$(wb.Sheets(j).Name) < UCase$(wb.Sheets(j + 1).Name) Then
            wb.Sheets(j).Move After:=wb.Sheets(j + 1)
        End If
    Next
Next
End Sub

Public Function returnFirstMatchFromCol(arr As Variant, lookupCol As Integer, returnCol As Integer, match As Variant) As String
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If CStr(arr(i, lookupCol) = CStr(match)) Then
            returnFirstMatchFromCol = arr(i, returnCol)
            Exit Function
        End If
    Next i
    returnFirstMatchFromCol = "NOT FOUND"
End Function
Public Function is2dArray(a As Variant) As Boolean
    If IsArrayEmpty(a) Then
        is2dArray = False
        Exit Function
    End If
    Dim l As Long
    On Error Resume Next
    l = LBound(a, 2)
    is2dArray = Err = 0
End Function
Public Function IsArrayEmpty(arr As Variant) As Boolean
Dim LB As Long, UB As Long
Err.Clear
On Error Resume Next
If IsArray(arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(arr, 1)
If (Err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBoung is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    LB = LBound(arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function
Public Function IsInString(stringToBeFound As String, stringList As String, Optional ByVal sep As String = ",") As Boolean
    Dim arr As Variant
    arr = Split(stringList, sep)
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInString = True
            Exit Function
        End If
    Next i
    IsInString = False
End Function
Public Function IsInArr(toBeFound As Variant, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = toBeFound Then
            IsInArr = True
            Exit Function
        End If
    Next i
    IsInArr = False
End Function
Public Function RowInArr(toBeFound As Variant, arr As Variant) As Long
    Dim i As Long
    ' arr is 1 row and many columns
    For i = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, i) = toBeFound Then
            Debug.Print (i)
            RowInArr = i
            Exit Function
        End If
    Next i
    RowInArr = 0
End Function
Function getNumbersFromString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .Pattern = "[^\d]+"
     getNumbersFromString = .Replace(strIn, vbNullString)
    End With
End Function

Function getArrFilteredToDates(ws As Worksheet, startd As Date, endD As Date, matchCol As Long) As Variant
    Dim lr As Long, lc As Long, farr As Variant
    With ws
        On Error Resume Next
            .ShowAllData
        On Error GoTo 0
        lr = .Range("A" & .Rows.Count).End(xlUp).Row ' gets last row with data
        lc = .Cells(1, .Columns.Count).End(xlToLeft).Column
        '1. Apply Filter
        .Range("A2", .Cells(lr, lc)).AutoFilter Field:=matchCol, Criteria1:=">=" & startd, Operator:=xlAnd, Criteria2:="<=" & endD
        '2. get filtered to Array
        farr = .Range("A2", .Cells(lr, lc)).SpecialCells(xlCellTypeVisible).Value2
        '3. Clear Filter
        On Error Resume Next
            .ShowAllData
        On Error GoTo 0
    End With
End Function
Function trimEmptyFromEndOfArray(arr As Variant, last As Long) As Variant
    Dim farr As Variant, i As Long, j As Long
    If last > 1 Then ' there is at least one
        ReDim farr(1 To last, 1 To UBound(arr, 2))
            For i = 1 To last ' for each row with stuff in it
                For j = 1 To UBound(arr, 2) ' for each column
                    farr(i, j) = arr(i, j) ' copy each field
                Next j
            Next i
        End If
    trimEmptyFromEndOfArray = farr
End Function
Public Function rotateArr(arr As Variant) As Variant
    ' takes in two 2D arrays and appends them
    Dim oarr As Variant, i As Long, j As Long
    ReDim oarr(1 To UBound(arr, 2), 1 To UBound(arr, 1)) ' make it big enough
    
    For i = 1 To UBound(arr, 1) ' for each row of orig array
        For j = 1 To UBound(arr, 2) ' for each column
            oarr(j, i) = arr(i, j)
        Next j
    Next i
    rotateArr = oarr
End Function
Public Function append2DArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    ' takes in two 2D arrays and appends them
    Dim totarr As Variant, i As Long, j As Long
    Debug.Print ("Arr1 " & UBound(arr1, 2) & " Arr2 " & UBound(arr2, 2))
    ReDim totarr(1 To UBound(arr1, 1) + UBound(arr2, 1), 1 To UBound(arr1, 2)) ' make it big enough
    
    For i = 1 To UBound(totarr, 1) ' for each row
        For j = 1 To UBound(totarr, 2) ' for each column
            If i <= UBound(arr1, 1) Then ' first array
                totarr(i, j) = arr1(i, j) ' first array
            Else
                totarr(i, j) = arr2(i - UBound(arr1, 1), j) ' second array
            End If
        Next j
    Next i
    Debug.Print ("Arr1 " & UBound(arr1, 1) & " Arr2 " & UBound(arr2, 1) & " TotArr " & UBound(totarr, 1))
   append2DArrays = totarr
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Outputs
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub outputArrinNew(arr)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    If is2dArray(arr) Then
        wb.Sheets(1).Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value2 = arr
    Else
        Call outputRowArrAtCell(wb.Sheets(1), "A1", arr)
    End If
End Sub
Sub outputArrAtCell(ws As Worksheet, cellStart As String, arr As Variant) ' for 2D Arrays
    Dim lx As Long, ly As Long 'x rows, y cols
    lx = UBound(arr, 1)
    ly = UBound(arr, 2)
    If (LBound(arr, 1) = 0) Then
        lx = lx + 1
    End If
    If (LBound(arr, 2) = 0) Then
        ly = ly + 1
    End If
    With ws
        .Range(cellStart).Resize(lx, ly).Value2 = arr
    End With
End Sub
Sub outputRowArrAtCell(ws As Worksheet, cellStart As String, arr As Variant) ' for 1D Arrays
    Dim lx As Long 'x rows, y cols
    lx = UBound(arr, 1)
    If (LBound(arr, 1) = 0) Then
        lx = lx + 1
    End If
    With ws
        .Range(cellStart).Resize(1, lx).Value2 = arr
    End With
End Sub
Sub appendArrToEndOfWS(ws As Worksheet, arr As Variant)
    Dim lr As Long
    With ws
        lr = .Range("A" & .Rows.Count).End(xlUp).Row ' gets last row with data
        .Range("A" & lr + 1).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Input Functions
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     WorksheetExists = Not sht Is Nothing
 End Function
 
 Public Function getDataToArray(wbkPath As String, wksIndex As Integer, Optional wksName As String = "NA", Optional pword As String = "NOTPROTECTED", Optional startCell As String = "A2") As Variant
    ' use this to get data out of an unopened workbook and then close it again
    Dim wbk As Workbook, wks As Worksheet, lastrow As Long, lastcol As Integer, ttarr As Variant
    If pword = "NOTPROTECTED" Then 'the file does not require a password
        Set wbk = Workbooks.Open(Filename:=wbkPath, ReadOnly:=True)
    Else
        Set wbk = Workbooks.Open(Filename:=wbkPath, ReadOnly:=True, Password:=pword)
    End If
    If wksName = "NA" Then 'no wksName was supplied, use index
        If wksIndex <> 0 Then
            Set wks = wbk.Sheets(wksIndex)
        Else
            Set wks = wbk.Sheets(1) 'assume first
        End If
    Else ' use wks index
        Set wks = wbk.Sheets(wksName)
    End If
    With wks
        lastrow = .Range(Left(startCell, 1) & .Rows.Count).End(xlUp).Row
        lastcol = .Cells(getNumbersFromString(startCell), .Columns.Count).End(xlToLeft).Column
        ttarr = .Range(startCell, .Cells(lastrow, lastcol)).Value2
        Debug.Print ("From " & wbkPath & "I have ttarr of size: " & lastrow & " and " & lastcol)
    End With
    wbk.Close SaveChanges:=False
    getDataToArray = ttarr
End Function
 Public Function getWKSDataToArray(wks As Worksheet, Optional startCell As String = "A2") As Variant
    ' this is a lot like the above function but used when you already have the workbook open
    Dim lastrow As Long, lastcol As Integer, ttarr As Variant
    With wks
        lastrow = .Range(Left(startCell, 1) & .Rows.Count).End(xlUp).Row
        lastcol = .Cells(getNumbersFromString(startCell), .Columns.Count).End(xlToLeft).Column
        ttarr = .Range(startCell, .Cells(lastrow, lastcol)).Value2
    End With
    getWKSDataToArray = ttarr
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Email Functions
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub sendEmailTo(recip As String, cc As String, subj As String, msg As String, Optional attach As String)
    ' this is called like so:
    ' Call BASE.sendEmailTo(recip:= recipstr, cc:= ccstr, subj:= substr, msg:= msgstr, attach:= pathstr)
    ' it opens MS Outlook and creates a message
    Dim objOutlook As Outlook.Application, objOutlookMsg As Outlook.MailItem, atts As Outlook.Attachments, newAtt As Outlook.Attachment
    Set objOutlook = CreateObject("Outlook.Application")
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
    objOutlookMsg.to = recip
    objOutlookMsg.cc = cc
    objOutlookMsg.Subject = subj
    objOutlookMsg.HTMLBody = msg
    If attach <> "" Then ' handle attachment
        Set atts = objOutlookMsg.Attachments
        Set newAtt = atts.Add(attach, olByValue)
    End If
    objOutlookMsg.Display
End Sub

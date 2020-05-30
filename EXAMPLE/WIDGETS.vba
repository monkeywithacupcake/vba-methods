
Option Explicit     ' all vars must be dimmed
Option Base 1       ' arrays start with 1
Option Compare Text ' That is, "AAA" is equal to "aaa".
'
'' WIDGETS MODULE
'
' The WIDGETS Module
' reads data from OTHER DIR/new.xlsx
' finds the corresponding file in Other Sub Dir
' creates a new tab for today with info from new
'

Public Sub callUpdateWithNewWidgets()
    ' this public sub can be called from a button
    ' the private sub below can only be seen inside this module
    Call updateWithNewWidgets
    ' done
End Sub
Private Sub updateWithNewWidgets()
    Call BASE.startProcessing
    ' get the new widget data
    ' new widget data are in a file (that you type into main)
    ' new widget data are of the form KEY COLOR LOGICAL NUMBER where KEY is the item that identifies the file we need to update
    Dim newWidgetTable As Variant, report As Variant
    newWidgetTable = BASE.getDataToArray(wbkPath:=LOOKUPS.getNewWidgetFilePath, wksIndex:=1)
    ' create a report array that will track what you did
    ReDim report(1 To UBound(newWidgetTable, 1), 1 To 2)
    ' for each line in the new widgets, process the new widget
    Dim i As Long, j As Long, k As String, arr As Variant
    ReDim arr(1 To 1, 1 To UBound(newWidgetTable, 2) - 1) ' this makes arr 1 row and as many columns as newwidgettable (less one for the key)
    For i = 1 To UBound(newWidgetTable, 1) ' for each row of widgets
        k = newWidgetTable(i, 1)
        For j = 1 To UBound(arr, 2)
            arr(1, j) = newWidgetTable(i, j + 1) ' we use j+1 because the first column in arr is the second column in new widgettable
        Next j
        Call processNewWidget(k, arr, report, i)
    Next i
    ' tell the user we are done
    BASE.outputArrinNew (report)
    MsgBox ("Your New Widgets Are Processed")
    Call BASE.endProcessing
End Sub

Private Sub processNewWidget(fkey As String, newInfo As Variant, reportArr As Variant, reportNum As Variant)
    ' this sub takes a widget key
    ' looks for a matching file in the widget directory
    Dim wbPath As String, wb As Workbook
    wbPath = LOOKUPS.getWidgetDir & Application.PathSeparator & fkey & ".xlsx"  ' change if .xls
    Debug.Print (wbPath)
    If Len(Dir(wbPath)) = 0 Then
        MsgBox ("Could not find the file for " & fkey)
    Else
        On Error Resume Next
        Set wb = Workbooks.Open(wbPath)
        On Error GoTo 0
        If wb Is Nothing Then MsgBox wbPath & " is invalid", vbCritical
    End If
    Debug.Print ("found a matching file for " & fkey)
    Call copyMostRecentAndAdd(wb:=wb, newInfo:=newInfo) ' copy and add
    reportArr(reportNum, 1) = Date
    reportArr(reportNum, 2) = "updated : " & fkey
    wb.Close SaveChanges:=True
End Sub

Private Sub copyMostRecentAndAdd(wb As Workbook, newInfo As Variant)
    Dim recent As String
    ' get the most recent
    recent = getMostRecentTab(wb)
    ' copy it
    wb.Sheets(recent).Copy After:=Sheets(recent)
    ' rename the new one
    wb.Sheets(recent & " (2)").Name = CStr(Date)
    ' add data
    Call BASE.appendArrToEndOfWS(wb.Sheets(CStr(Date)), newInfo)
End Sub

Private Function getMostRecentTab(wb As Workbook) As String
    Dim i As Long, tabList As Variant, dValue As Long
    
    ReDim tabList(1 To wb.Sheets.Count)
    For i = 1 To wb.Sheets.Count
    On Error Resume Next
        dValue = DateValue(wb.Sheets(i).Name)
        If IsError(dValue) Or IsEmpty(dValue) Then   ' do nothing
        Else
            tabList(i) = CLng(dValue)
        End If
    Next i
    
    getMostRecentTab = CStr(CDate(WorksheetFunction.Max(tabList)))
On Error GoTo 0
End Function


Option Explicit     ' all vars must be dimmed
Option Base 1       ' arrays start with 1
Option Compare Text ' That is, "AAA" is equal to "aaa".
'
'' LOOKUPS Module

' LOOKUPS Module is specific to this file
' It contains functions that return paths and values to make this file work
' Some functions will be reused over projects
' Most functions in a LOOKUPS module will be Public

Private Sub whydeclare()
    ' Put your mouse here and click the "play" button above.
    ' You should see the following two sentences in the "Immediate Window"
    ' Don't have an Immediate Window below? Click 'View' and select it
    Debug.Print ("Declare every Sub and Function Public or Private as a matter of practice")
    Debug.Print ("Not sure? It should be Private")
    ' Private subs and functions cannot be used outside of the module in which they are declared
End Sub

Public Function getMainDir() As String ' declare what your function will return
    Dim str As String ' gonna use a variable? declare it
    str = ThisWorkbook.Names("DIR_MAIN").RefersToRange.Value2  ' remember our named ranges on the "MAIN" tab?
    ' what do you think is in str right now?
    ' if this file is in this folder, you can also use: ThisWorkbook.Path
    Debug.Print (str) ' you can see in the Immediate Window
    getMainDir = str   ' you have to write function output/return lines like this
End Function

Public Function getSampleDir() As String
    getSampleDir = ThisWorkbook.Names("DIR_SAMPLE").RefersToRange.Value2 ' you can do the whole thing in one line
End Function

Public Function getOtherDir() As String
    getOtherDir = getMainDir() & Application.PathSeparator & ThisWorkbook.Names("DIR_OTHER").RefersToRange.Value2 ' you can do the whole thing in one line
End Function


Public Function getWidgetDir() As String
    ' this is where the excel files for each widget are located
    getWidgetDir = getMainDir() & Application.PathSeparator & ThisWorkbook.Names("DIR_WIDGETS").RefersToRange.Value2
End Function
Public Function getNewWidgetFilePath() As String
    ' this is the Excel file where we have new data for our widgets
    getNewWidgetFilePath = getMainDir() & Application.PathSeparator & ThisWorkbook.Names("FILE_NEW_WIDGETS").RefersToRange.Value2
End Function

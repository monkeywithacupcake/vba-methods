Sub createPDFFromWordTemplate(pCount As Integer, t1Rng As Range, t2Rng As Range, wTmp As String, pOut As String)
    'this will NOT work out of the box, you must update
    'wTmp is a path to a Word Template FILE
    'pOut is a path to a pdf Output DIRECTORY
    'assumption here that we want a number 'pCount' and we want two tables from a worksheet (t1Rng, t2Rng) 
    'the number and table will be placed at a bookmark in a Microsoft Word Template
    Dim wdApp As Object, wd As Object

    Set wdApp = CreateObject("Word.Application")
    Set wd = wdApp.Documents.Add(Template:=wTmp, NewTemplate:=False, DocumentType:=wdNewBlankDocument) 'creates new from template
    wdApp.Visible = True

    With wd
       .Bookmarks("desc_Tab1").Range.Text = pCount & " " 'assumes this bookmark exists in template
       t1Rng.Copy
       .Bookmarks("Tab1").Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=True
       t2Rng.Copy
       .Bookmarks("Tab2").Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=True
       
       Dim i As Integer
       For i = 1 To 2 'format both of the tables
           .Tables(i).Style = "Plain Table 3"
           .Tables(i).AutoFitBehavior wdAutoFitWindow
           .Tables(i).Range.Font.Size = 10
           .Tables(i).Range.Font.name = "Calibri"
           .Tables(i).Columns.AutoFit
           .Tables(i).Rows.SetHeight 13, 2
           .Tables(i).Rows(1).HeadingFormat = True
       Next i
   
       .SaveAs (pOut & "newDocument" & Format(Now, "YYYYMMDD") & ".pdf"), 17
       .Close SaveChanges:=False
   
    End With
    wdApp.Quit
End Sub

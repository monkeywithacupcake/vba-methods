Sub sendOutlookEmail(recip As String, subj As String, msg As String, Optional attach As String)
    ' Will open Outlook with a Draft Message that is To: recip, with Subject: subj, Body: msg, and optional Attachment
    ' Caution: you must have a Reference to Microsoft Outlook 16.0 Object Library 
    ' In Excel, click "Developer" on the ribbon, click "Tools", click "References", and check the box  Microsoft Outlook 16.0 Object Library
    
    Dim objOutlook As Outlook.Application, objOutlookMsg As Outlook.MailItem, atts As Outlook.Attachments, newAtt As Outlook.Attachment

    Set objOutlook = CreateObject("Outlook.Application")
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

    objOutlookMsg.To = recip
    objOutlookMsg.Subject = subj
    objOutlookMsg.HTMLBody = msg
    
    If attach <> "" Then ' handle attachment
        Set atts = objOutlookMsg.Attachments
        Set newAtt = atts.Add(attach, olByValue)
    End If
    objOutlookMsg.Display 
End Sub

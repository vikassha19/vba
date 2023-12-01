Sub SaveAttachmentAndAppendData(Item As Outlook.MailItem)
    Dim attachment As Outlook.Attachment
    Dim excelApp As Object
    Dim workbook As Object
    Dim targetWorkbook As Object
    
    ' Check if the email is unread (not processed)
    If Not Item.UnRead Then Exit Sub
    
    ' Check if the email has attachments
    If Item.Attachments.Count > 0 Then
        ' Get the attachment
        Set attachment = Item.Attachments(1) ' Assuming there's only one attachment
        ' Save the attachment to a folder
        attachment.SaveAsFile "C:\Path\To\Save\Previous_Day_INC_Ticket_Report.xlsx"
        
        ' Mark the email as read (processed)
        Item.UnRead = False
        
        ' Excel setup
        Set excelApp = CreateObject("Excel.Application")
        Set workbook = excelApp.Workbooks.Open("C:\Path\To\Save\Previous_Day_INC_Ticket_Report.xlsx")
        
        ' Append data to another workbook
        Set targetWorkbook = excelApp.Workbooks.Open("C:\Path\To\Your\Target\Workbook.xlsx")
        ' Customize the following line based on your data structure
        ' For example, assuming data is in the first sheet of both workbooks
        workbook.Sheets(1).UsedRange.Copy targetWorkbook.Sheets(1).Cells(targetWorkbook.Sheets(1).Rows.Count, 1).End(-4162)
        
        ' Save and close workbooks
        targetWorkbook.Save
        targetWorkbook.Close
        workbook.Close
        excelApp.Quit
    End If
    
    ' Clean up
    Set excelApp = Nothing
End Sub

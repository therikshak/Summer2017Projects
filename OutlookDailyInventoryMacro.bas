Attribute VB_Name = "CreateInventoryReportOutlook"
Public Sub completeDailyInventory()
    'delete everything in sharepoint folder except Product Information
    DeleteReports
    
    'download reports from emails and save to sharepoint folder
    Dim successfulDownload As Boolean
    successfulDownload = DownloadReports
    
    'Run Script to get Lindner inventory
        'Dim wsh As Object
        'Set wsh = CreateObject("Wscript.Shell")
        'path = "C:\Users\estryshak\SharePoint\T\Projects\InventoryReports\Macro\LindnerScript.exe"
        'Call wsh.Run(path, 0, True)
    
    'create new excel workbook in sharepoint folder and run Daily Inventory Macro
    'to create the pivot table
    If successfulDownload Then
        CreatePivot
    End If
    
    'move old reports
    moveOld
End Sub


'delete the old reports
Private Sub DeleteReports()
    Dim path As String
    Dim username As String
    username = (Environ$("Username"))
    'path to the folder
    path = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    
    Dim todayDate As Date, m As String, d As String, y As String, combinedDate As String
    todayDate = DateValue(Date)
    m = month(todayDate)
    d = day(todayDate)
    y = year(todayDate)
    combinedDate = (m & "_" & d & "_" & y)

    file = Dir(path)
    'loop through all files in folder
    Do While Len(file) > 0
        If InStr(1, file, "ProductInformation") > 0 Then
            'skip the product information file
        ElseIf InStr(1, file, combinedDate) > 0 Then
            'skip deleting report if it is today's and macro was run again
        Else
            'delete the file
            Kill path & file
        End If
        'get next file
        file = Dir
    Loop
    
End Sub

Private Function DownloadReports() As Boolean
    Dim Item As Outlook.MailItem
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.folder
    Dim reportFolder As Outlook.folder

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set reportFolder = myInbox.Folders("Inventory Reports Macro")
    
    'subject and sender of email and variable to store attachments in
    Dim subject As String, sender As String, attachments As Outlook.attachments
    
    'number of reports downloaded, number of attachments for an email, number of city reports downloaded
    Dim numDownloaded As Integer, attachmentCount As Long, cityCount As Integer
    
    'i used to count through attachments, fileNameLength is helper variable to determine file type, theFileType
    'gets the fileType of the attachment
    Dim i As Integer, fileNameLength As Long, theFileType As String
    
    'saveFolder is where reports will be saved, username gets the system username, fileName gets the report name
    'filepath will have the folder, name of the report, and other information concatenated to save the report
    Dim saveFolder As String, username As String, fileName As String, filepath As String
    
    username = (Environ$("Username"))
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    
    'cityCount counts how many city brewery reports have been saved so it can make each filename unique
    cityCount = 1
    'count how many reports downloaded
    numDownloaded = 0
    
    For Each Item In reportFolder.items
        'if it is an email, then get its data
        If Item.Class = olMail Then
            'GET EMAIL INFO
            'subject line of email
            subject = Item.subject
            'original sender
            sender = Item.SenderEmailAddress
                        
            'IF FROM NEW HOLLAND
            If sender = "payables@newhollandbrew.com" Then
                'report comes as text in the body of the email
                'so it needs to be put into an excel file
                exportToExcel Item, saveFolder
                numDownloaded = numDownloaded + 1
                newHoll = True
            'if not New Holland then get the attachments
            Else
                'SAVE ATTACHMENTS
                Set attachments = Item.attachments
                attachmentCount = attachments.Count
                'if attachments exist
                For i = attachmentCount To 1 Step -1
                    'get first filename attached to email
                    filepath = attachments.Item(i).fileName
                    'get the length up until a period
                    fileNameLength = InStrRev(filepath, ".")
                    'make a substring that is the fileType and the fileName without type
                    theFileType = Right(filepath, Len(filepath) - fileNameLength)
                    fileName = Left(filepath, fileNameLength - 1)
                    'only save the attachment if it is an excel file
                    If theFileType = "xls" Or theFileType = "xlsx" Then
                        'if a city brewery, then a number needs to be added to the end of the file
                        If InStr(1, filepath, "AGED FG") > 0 Then
                            filepath = saveFolder & fileName & cityCount & "." & theFileType
                            cityCount = cityCount + 1
                        Else
                            filepath = saveFolder & filepath
                        End If
                        'save the attachment to the specified location
                        attachments.Item(i).SaveAsFile filepath
                    End If
                Next i
                'increment number of reports downloaded
                numDownloaded = numDownloaded + 1
            End If
        End If
    Next Item

    Set Item = Nothing
    Set reportFolder = Nothing
    
    'set boolean for if reports downloaded
    If numDownloaded > 0 Then
        DownloadReports = True
        Exit Function
    Else
        DownloadReports = False
        Exit Function
    End If
    
End Function

'export email contents to excel
'takes in a mail item and where to save the file
Private Sub exportToExcel(mail As Outlook.MailItem, folder As String)
    Dim fileName As String, filepath As String
    fileName = "NewHollandReport.xlsx"
    filepath = folder & fileName
    
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim lRow As Long
    
    'create instance of excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    
    Dim tableRows() As String, tableCols() As String, destCell As Object
    Dim r As Integer, C As Integer
    Set destCell = xlWs.Range("A1")
    'get the rows of the table in the email
    tableRows = Split(mail.Body, vbCrLf)
    'loop through each row
    For r = 2 To UBound(tableRows)
        'read if there are empty cells after useful data
        If Len(tableRows(r)) < 5 Then Exit For
        'get each cell and put into excel
        tableCols = Split(tableRows(r), vbTab)
        For C = 0 To UBound(tableCols)
            destCell.Offset(r - 2, C).Value = tableCols(C)
        Next
    Next
    
    'save the excel file
    xlWb.Application.DisplayAlerts = False
    Application.DisplayAlerts = False
    
    xlWb.SaveAs filepath
    
    xlWb.Application.DisplayAlerts = True
    Application.DisplayAlerts = True
    
    xlWb.Close
    xlApp.Quit
    
    Set xlApp = Nothing
    Set xlWb = Nothing
    Set xlWs = Nothing
End Sub

Private Sub moveOld()
    Dim Item As Outlook.MailItem 'used for individual emails
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.folder
    Dim myDestFolder As Outlook.folder
    Dim reportFolder As Outlook.folder

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myDestFolder = myInbox.Folders("Old Inventory Reports")
    Set reportFolder = myInbox.Folders("Inventory Reports Macro")
    
    While reportFolder.items.Count > 0
        For Each Item In reportFolder.items
            If Item.Class = olMail Then
                Item.Move myDestFolder
            End If
        Next Item
    Wend

End Sub

'create the inventory report
Private Sub CreatePivot()
    Dim xlApp As Excel.Application
    Dim xlWb As Workbook, xlWs As Object, wlWb2 As Workbook
    Dim username As String
    username = (Environ$("Username"))
    
    'start a new instance of excel
    Set xlApp = New Excel.Application
    Set xlwb2 = xlApp.Workbooks.Open("C:\Users\" & username & "\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.xlsb")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    
    'get today's date
    Dim todayDate As Date, m As String, d As String, y As String
    todayDate = DateValue(Date)
    m = month(todayDate)
    d = day(todayDate)
    y = year(todayDate)
    
    'filename for inventory report to be saved as
    Dim fileName As String
    fileName = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\" & m & "_" & d & "_" & y & "_" & "InventoryReport.xlsx"

    'don't visibly open excel
    xlApp.Visible = False
    'run the inventory macro from the PERSONAL workbook
    xlWb.Application.Run "PERSONAL.XLSB!DailyInventory.DailyInventory"
    
    'save the excel file and close
    xlWb.Application.DisplayAlerts = False
    Application.DisplayAlerts = False
    xlWb.SaveAs fileName
    xlWb.Application.DisplayAlerts = True
    Application.DisplayAlerts = True                                                                                       
                                                                                            
    xlWb.Close
    xlwb2.Close
    xlApp.Quit
    
    Set xlApp = Nothing
    Set xlWb = Nothing
    Set xlWs = Nothing
    Set xlwb2 = Nothing
End Sub

Private Sub Application_Startup()
  'CreateAppointment
End Sub

Private Sub Application_Reminder(ByVal Item As Object)
Set olRemind = Outlook.reminders

If Item.MessageClass <> "IPM.Appointment" Then
  Exit Sub
End If
 
If Item.Categories <> "Run in 5" Then
  Exit Sub
End If
 
' Call your macro here
completeDailyInventory

'Delete Appt from calendar when finished
Item.Delete

' Create another appt to repeat the process
' CreateAppointment

End Sub

' dismiss reminder
Private Sub olRemind_BeforeReminderShow(Cancel As Boolean)

    For Each objRem In olRemind
            If objRem.Caption = "This Appointment reminder fires in 5" Then
                If objRem.IsVisible Then
                    objRem.Dismiss
                    Cancel = True
                End If
                Exit For
            End If
        Next objRem
End Sub

' Put this macro in a Module
Public Sub CreateAppointment()
Dim objAppointment As Outlook.AppointmentItem
Dim tDate As Date
' Using a 1 min reminder so 6  = reminder fires at 5 min.
tDate = Now() + 2 / 1440

Set objAppointment = Application.CreateItem(olAppointmentItem)
      With objAppointment
        .Categories = "Run in 5"
        .Body = "This Appointment reminder fires in 5"
        .Start = tDate
        .End = tDate
        .subject = "This Appointment reminder fires in 5"
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 1
        .Save
      End With
End Sub

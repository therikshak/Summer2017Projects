Attribute VB_Name = "CreateInventoryReportOutlook"
Public Sub completeDailyInventory()
    '********************* SET UP LOGGING AND SAVE FOLDER ************************
    Dim username As String, saveFolder As String, logPath As String, TextFile As Integer
    username = (Environ$("Username"))
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    logPath = saveFolder & "logGeneral.txt"
    TextFile = FreeFile
    Open logPath For Output As TextFile
    Print #TextFile, Now
    
    '*********************** Report Run *********************************
    Dim alreadyRan As Boolean
    Print #TextFile, "checking to see if report has been run"
    Dim question As String, answer As Variant
    alreadyRan = ReportRun
    If alreadyRan Then
        question = "It looks like the Inventory Report has been run today, would you like to rerun it?"
        answer = MsgBox(question, vbYesNo)
        Select Case answer
        Case vbYes
            'continue with execution
            Print #TextFile, "user elected to rerun"
        Case vbNo
            Print #TextFile, "user elected to not rerun"
            Close TextFile
            Exit Sub
        End Select
    End If
    '********************* DELETE FILES IN FOLDER ************************
    Print #TextFile, "Before DeleteReports"
    'delete everything in sharepoint folder except Product Information
    DeleteReports
    Print #TextFile, "After DeleteReports"
    
    '******************* DOWNLOAD REPORTS FROM EMAILS *********************
    Print #TextFile, "Before DownloadReports"
    'download reports from emails and save to sharepoint folder
    Dim successfulDownload As Boolean
    successfulDownload = DownloadReports
    If Not successfulDownload Then
        Print #TextFile, "did not download any reports"
        GoTo move_old_skip
    End If
    
    '***************************** LINDNER ********************************
    Print #TextFile, "Before Lindner Call"
    'Run Script to get Lindner inventory, this saves to Documents
    Dim wsh As Object, x As Integer
    Set wsh = VBA.CreateObject("Wscript.Shell")
    On Error GoTo failed_lindner
    x = wsh.Run("cmd /c C:\Users\" & username & "\Desktop\Lindner_Scrape\build\exe.win32-3.6\Scrape.exe", 1, True)
    Set wsh = Nothing
    moveLinder
failed_lindner:
    Print #TextFile, "After Lindner Call"
    
    On Error GoTo 0
    '************************** CREATE TABLE ******************************
    'create new excel workbook in sharepoint folder and run Daily Inventory Macro
    'to create the pivot table
    Dim successful_pivot As Boolean
    successful_pivot = CreateTable
    If successful_pivot Then
        Print #TextFile, "CreateTable successful"
    Else
        Print #TextFile, "CreateTable not successful"
        Close TextFile
        Exit Sub
    End If
      
    '******************* MOVE OLD REPORT EMAILS ********************
move_old_skip:
    moveOld
    Print #TextFile, "moveOld called"
    Close TextFile
End Sub

' check if report has been run already
Private Function ReportRun() As Boolean
    Dim path As String
    Dim username As String, file As Variant
    username = (Environ$("Username"))
    'path to the folder
    path = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    ' get todays date
    Dim todayDate As Date, m As String, d As String, y As String, combinedDate As String
    todayDate = DateValue(Date)
    m = month(todayDate)
    d = day(todayDate)
    combinedDate = (m & "_" & d)
    
    file = Dir(path)
    'loop through all files in folder
    Do While Len(file) > 0
        If InStr(1, file, combinedDate) > 0 Then
            'exit sub and pass true back to the variable already ran
            ReportRun = True
            Exit Function
        Else
            'skip file
        End If
        'get next file
        file = Dir
    Loop
    ' pass false back to the variable already ran
    ReportRun = False
End Function

'delete the old reports
Private Sub DeleteReports()
    Dim path As String
    Dim username As String, file As Variant
    username = (Environ$("Username"))
    'path to the folder
    path = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"

    file = Dir(path)
    'loop through all files in folder
    Do While Len(file) > 0
        If InStr(1, file, "ProductInformation") > 0 Then
            'skip the product information file
        ElseIf InStr(1, file, "General") > 0 Then
            ' skip deleting the general log
        Else
            'delete the file
            Kill path & file
        End If
        'get next file
        file = Dir
    Loop
End Sub

' move the lindner csv to the sharepoint folder
Private Sub moveLinder()
    Dim username As String, lindner_original_location As String, saveFolder As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    ' get the path to my documents and the sharepoint folder
    username = (Environ$("Username"))
    lindner_original_location = "C:\Users\" & username & "\Desktop\Lindner_Scrape\lindner.csv"
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    
    ' move from documents to the sharepoint folder
    FSO.MoveFile Source:=lindner_original_location, Destination:=saveFolder

End Sub

' Download all of the reports in the folder and save to the sharepoint folder
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
    
    Dim recMonth As String, recDay As String, dateReceived As String
    
    'number of reports downloaded, number of attachments for an email, number of city reports downloaded
    Dim numDownloaded As Integer, attachmentCount As Long, cityCount As Integer
    
    'i used to count through attachments, fileNameLength is helper variable to determine file type, theFileType
    'gets the fileType of the attachment
    Dim i As Integer, fileNameLength As Long, theFileType As String
    
    'saveFolder is where reports will be saved, username gets the system username, fileName gets the report name
    'filepath will have the folder, name of the report, and other information concatenated to save the report
    Dim saveFolder As String, username As String, fileName As String, filepath As String, documentsFolder As String

    username = (Environ$("Username"))
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    
    'cityCount counts how many city brewery reports have been saved so it can make each filename unique
    cityCount = 1
    'count how many reports downloaded
    numDownloaded = 0
    
    Dim logPath As String, TextFile As Integer
    logPath = saveFolder & "logDownloadReports.txt"
    
    ' create a new log, the file will contain:
        ' Date & time report run
        ' skipped & saved emails by sender
        ' number of saved reports
        
    TextFile = FreeFile
    Open logPath For Output As TextFile
    Print #TextFile, Now
    
    For Each Item In reportFolder.items
        'if it is an email, then get its data
        If Item.Class = olMail Then
            ' GET EMAIL INFO: subject, sender, date received
            subject = Item.subject
            sender = Item.SenderEmailAddress
            dateReceived = Item.ReceivedTime
            recMonth = month(dateReceived)
            recDay = day(dateReceived)
            
            ' IF FROM NEW HOLLAND
            If sender = "payables@newhollandbrew.com" Then
            ' check if this is the correct email and export to excel if so otherwise the email will be skipped
                If (goodEmail(recDay, recMonth, True)) Then
                    ' report comes as text in the body of the email so it needs to be put into an excel file
                    exportToExcel Item, saveFolder
                    Print #TextFile, "saved: " & sender
                    numDownloaded = numDownloaded + 1
                Else
                    Print #TextFile, "skipped: " & sender
                End If
            ' if not New Holland then get the attachments
            Else
                ' first check if the email is from today, otherwise skip it
                If (goodEmail(recDay, recMonth, False)) Then
                    ' SAVE ATTACHMENTS
                    Set attachments = Item.attachments
                    attachmentCount = attachments.Count
                    ' if attachments exist
                    For i = attachmentCount To 1 Step -1
                        ' get first filename attached to email
                        filepath = attachments.Item(i).fileName
                        ' get the length up until a period
                        fileNameLength = InStrRev(filepath, ".")
                        ' make a substring that is the fileType and the fileName without type
                        theFileType = Right(filepath, Len(filepath) - fileNameLength)
                        fileName = Left(filepath, fileNameLength - 1)
                        ' only save the attachment if it is an excel file
                        If theFileType = "xls" Or theFileType = "xlsx" Then
                            ' if a city brewery, then a number needs to be added to the end of the file
                            If InStr(1, filepath, "AGED FG") > 0 Then
                                filepath = saveFolder & fileName & cityCount & "." & theFileType
                                cityCount = cityCount + 1
                            Else
                                filepath = saveFolder & filepath
                            End If
                            ' save the attachment to the specified location
                            attachments.Item(i).SaveAsFile filepath
                            Print #TextFile, "saved: " & sender
                        End If
                    Next i
                    ' increment number of reports downloaded
                    numDownloaded = numDownloaded + 1
                Else
                    Print #TextFile, "skipped: " & sender
                End If
            End If
        End If
    Next Item

    Set Item = Nothing
    Set reportFolder = Nothing
    
    Print #TextFile, numDownloaded
    Close TextFile
    
    ' set boolean for if reports downloaded
    If numDownloaded > 0 Then
        DownloadReports = True
        Exit Function
    Else
        DownloadReports = False
        Exit Function
    End If
    
End Function

' takes in the day and month of the received email as well as if it is from new holland
' if the email was received today, it returns true (for New Holland it is yesterday for true)
Private Function goodEmail(recDay As String, recMonth As String, newHoll As Boolean) As Boolean
    If newHoll Then
        If recDay <> day(Date - 1) Then
            GoTo wrongDate
            If recMonth <> month(Date - 1) Then
                GoTo wrongDate
            End If
        End If
    Else
        If recDay <> day(Date) Then
            GoTo wrongDate
            If recMonth <> month(Date) Then
                GoTo wrongDate
            End If
        End If
    End If
    goodEmail = True
    Exit Function
wrongDate:
    goodEmail = False
End Function

' export email contents to excel
' takes in a mail item and where to save the file
Private Sub exportToExcel(mail As Outlook.MailItem, folder As String)
    Dim fileName As String, filepath As String
    fileName = "NewHollandReport.xlsx"
    filepath = folder & fileName
    
    ' log what is happening in logExportToExcel.txt
    Dim username As String, saveFolder As String, logPath As String, TextFile As Integer
    username = (Environ$("Username"))
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    logPath = saveFolder & "logExportToExcel.txt"
    TextFile = FreeFile
    Open logPath For Output As TextFile
    Print #TextFile, Now
    
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim lRow As Long
    
    'create instance of excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    
    Print #TextFile, "Created instance of excel"
    
    Dim tableRows() As String, tableCols() As String, destCell As Object
    Dim r As Integer, C As Integer
    Set destCell = xlWs.Range("A1")
    'get the rows of the table in the email
    tableRows = Split(mail.Body, vbCrLf)
    Print #TextFile, "Before Looping"
    'loop through each row
    For r = 2 To UBound(tableRows)
        ' read if there are empty cells and exit if there are
        If Len(tableRows(r)) < 5 Then Exit For
        ' get each cell and put into excel
        tableCols = Split(tableRows(r), vbTab)
        For C = 0 To UBound(tableCols)
            destCell.Offset(r - 2, C).Value = tableCols(C)
        Next
    Next
    
    Print #TextFile, "After looping through email"
    'save the excel file and close excel
    xlWb.Application.DisplayAlerts = False
    xlWb.SaveAs filepath
    xlWb.Application.DisplayAlerts = True
    Print #TextFile, "After saving excel"
    xlWb.Close
    xlApp.Quit
    Print #TextFile, "After closing workbook and quitting the application"
    
    Set xlApp = Nothing
    Set xlWb = Nothing
    Set xlWs = Nothing
    Close TextFile
End Sub

'create the inventory report
Private Function CreateTable() As Boolean
    Dim xlApp As Excel.Application
    Dim xlWb As Workbook, xlWs As Object, xlWb2 As Workbook
    Dim username As String
    username = (Environ$("Username"))
    
    ' LOGGING
    Dim saveFolder As String, logPath As String, TextFile As Integer
    saveFolder = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    logPath = saveFolder & "logCreateTable.txt"
    TextFile = FreeFile
    Open logPath For Output As TextFile
    Print #TextFile, Now
    
    'start a new instance of excel
    Set xlApp = New Excel.Application
    On Error GoTo noPersonal
    Set xlWb2 = xlApp.Workbooks.Open("C:\Users\" & username & "\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.xlsb")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    Print #TextFile, "Created excel app and opened PERSONAL"
    
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
    On Error GoTo failedMacro
    xlWb.Application.Run "PERSONAL.XLSB!DailyInventory"
    Print #TextFile, "Macro Run and finished"
    
    'save the excel file and close
    xlWb.SaveAs fileName
    xlWb.Close
    xlWb2.Close
    xlApp.Quit
    
    Set xlApp = Nothing
    Set xlWb = Nothing
    Set xlWs = Nothing
    Set xlWb2 = Nothing
    
    Print #TextFile, "Excel saved, closed, and quit"
    Close TextFile
    CreateTable = True
    Exit Function
    
noPersonal:
    Set xlApp = Nothing
    Print #TextFile, "noPersonal error raised"
    Close TextFile
    CreateTable = False
    Exit Function
    
failedMacro:
    xlWb.Close False
    xlWb2.Close False
    xlApp.Quit
    Set xlApp = Nothing
    Set xlWb = Nothing
    Set xlWs = Nothing
    Set xlWb2 = Nothing
    Print #TextFile, "failedMacro error raised"
    Close TextFile
    CreateTable = False
    
End Function

' move all of the old reports to the old reports folder
Private Sub moveOld()
    Dim Item As Outlook.MailItem 'used for individual emails
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.folder
    Dim myDestFolder As Outlook.folder
    Dim reportFolder As Outlook.folder
    Dim recMonth As String, recDay As String, dateReceived As String

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myDestFolder = myInbox.Folders("Old Inventory Reports")
    Set reportFolder = myInbox.Folders("Inventory Reports Macro")
    
    Dim iteration_count As Integer
    iteration_count = 0
    ' loop through each email in the Inventory Reports Macro folder
    ' and move yesterday's to the old folder
    For Each Item In reportFolder.items
        If Item.Class = olMail Then
            dateReceived = Item.ReceivedTime
            recMonth = month(dateReceived)
            recDay = day(dateReceived)
        ' IF FROM NEW HOLLAND
            If sender = "payables@newhollandbrew.com" Then
            ' check if this is the correct email and export to excel if so otherwise the email will be skipped
                If Not (goodEmail(recDay, recMonth, True)) Then
                    Item.Move myDestFolder
                End If
            Else
                ' first check if the email is not from today, otherwise skip it
                If Not (goodEmail(recDay, recMonth, False)) Then
                    Item.Move myDestFolder
                End If
            End If
        End If
        iteration_count = iteration_count + 1
    Next Item
End Sub

'On start up create an appointment to trigger the inventory report to run
Private Sub Application_Startup()
  'CreateAppointment
End Sub

' If a reminder pops up check it and if it is the "Run Inventory" macro, then run the inventory report
' otherwise exit
Private Sub Application_Reminder(ByVal Item As Object)
    Set olRemind = Outlook.reminders
    If Item.MessageClass <> "IPM.Appointment" Then
      Exit Sub
    End If
     
    If Item.Categories <> "Run Inventory" Then
      Exit Sub
    End If
     
    ' Call macro to run
    completeDailyInventory
    
    'Delete Appt from calendar when finished
    Item.Delete
End Sub

' dismiss reminder
Private Sub olRemind_BeforeReminderShow(Cancel As Boolean)
    ' check each reminder and if it is the Run Inventory reminder, then dismiss it
    For Each objRem In olRemind
        If objRem.Caption = "Run Inventory" Then
            If objRem.IsVisible Then
                objRem.Dismiss
                Cancel = True
            End If
            Exit For
        End If
    Next objRem
End Sub

' create an appointment that will trigger one minute from now
Public Sub CreateAppointment()
    Dim objAppointment As Outlook.AppointmentItem
    Dim tDate As Date
    tDate = Now() + 2 / 1440
    
    Set objAppointment = Application.CreateItem(olAppointmentItem)
          With objAppointment
            .Categories = "Run Inventory"
            .Body = "Run Inventory"
            .Start = tDate
            .End = tDate
            .subject = "Run Inventory"
            .ReminderSet = True
            .ReminderMinutesBeforeStart = 1
            .Save
          End With
End Sub

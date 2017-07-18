Attribute VB_Name = "NetInventoryMacro"
Public Sub NetInventory()
    ' put the following excel files into this folder
        ' daily inventory report
        ' transfer orders
        ' purchase orders'
        ' vbs copy paste
        
    ' store path to folder containing files
    Application.ScreenUpdating = False
    Dim username As String, path As String
    username = (Environ$("Username"))
    path = "C:\Users\" & username & "\Desktop\AX_Export\"
    
    ' Setup variables for different workbooks and sheets
    Dim shtMasterModesto As Worksheet, shtMasterJoliet As Worksheet
    Dim wkbInventory As Workbook, shtInventory As Worksheet
    Dim wkbTransferOrder As Workbook, shtTransferOrder As Worksheet
    Dim wkbPurchaseOrder As Workbook, shtPurchaseOrder As Worksheet
    Dim wkbVbs As Workbook, shtVbs As Worksheet
    
    ' Assign sheet for Modesto and Joliet
    Sheets.Add Before:=ActiveSheet
    ActiveSheet.Name = "Modesto"
    Set shtMasterModesto = Worksheets("Modesto")
    Set shtMasterJoliet = Worksheets(2)
    shtMasterJoliet.Name = "Joliet"
    
    ' in this report, column headers will be:
    '   A   B    C       D           E                 F            G  H       I           J
    ' Plant|AX|Prod8|Description|Quantity(vbs)|Inventory|PO|TO|Total_projected|Diff
        ' total_projected = TO + PO + quantity(inv report)
        ' diff = total_projected - quanity(vbs)
    
    shtMasterModesto.Range("A1") = "Plant"
    shtMasterModesto.Range("B1") = "AX #"
    shtMasterModesto.Range("C1") = "Prod 8"
    shtMasterModesto.Range("D1") = "Description"
    shtMasterModesto.Range("E1") = "Quantity(vbs)"
    shtMasterModesto.Range("F1") = "Inventory"
    shtMasterModesto.Range("G1") = "PO"
    shtMasterModesto.Range("H1") = "TO"
    shtMasterModesto.Range("I1") = "Total_Projected"
    shtMasterModesto.Range("J1") = "Difference"
    
    shtMasterJoliet.Range("A1") = "Plant"
    shtMasterJoliet.Range("B1") = "AX #"
    shtMasterJoliet.Range("C1") = "Prod 8"
    shtMasterJoliet.Range("D1") = "Description"
    shtMasterJoliet.Range("E1") = "Quantity(vbs)"
    shtMasterJoliet.Range("F1") = "Inventory"
    shtMasterJoliet.Range("G1") = "PO"
    shtMasterJoliet.Range("H1") = "TO"
    shtMasterJoliet.Range("I1") = "Total_Projected"
    shtMasterJoliet.Range("J1") = "Difference"

' *************************************************************************************************************
    ' Step 1: Move data from vbs to this sheet
        ' leave off barrels
        ' range A2:D sheet.Cells(Rows.Count, 1).End(xlUp).Row
    Set wkbVbs = Workbooks.Open(path & "vbs_pull.xlsx")
    Set shtVbs = wkbVbs.Sheets(1)
    Dim numRows As Integer
    
    numRows = shtVbs.Cells(Rows.Count, 2).End(xlUp).Row
    ' get rid of formatting
    With shtVbs.Columns("A:E")
        .WrapText = False
        .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    shtVbs.Range("B2:D" & numRows).Copy Destination:=shtMasterModesto.Cells(2, "C")
    shtVbs.Range("B2:D" & numRows).Copy Destination:=shtMasterJoliet.Cells(2, "C")
    
    'close the workbook
    wkbVbs.Close (False)

' *************************************************************************************************************
    ' Step 2: get the inventory totals from daily report
        ' first filter the report to the correct brewery, then vlookup
        ' below goes into column F:
        ' iferror(vlookup(B#,dailyinventory sheet( Range C2:D end of table), 2 (units), 0) , value if true = 0)
        ' after getting units, do an index match to put ax numbers in
    Dim lastRow
    Dim invReportName As String
    Dim todayDate As Date, m As String, d As String, y As String
    todayDate = DateValue(Date)
    m = Month(todayDate)
    d = Day(todayDate)
    y = Year(todayDate)
    invReportName = m & "_" & d & "_" & y & "_InventoryReport"
    
    Set wkbInventory = Workbooks.Open(path & invReportName)
    Set shtInventory = wkbInventory.Sheets("Daily Inventory")
    
    ' filter by joliet and perform vlookup
    shtInventory.ListObjects("Table2").Range.AutoFilter Field:=1, Criteria1:="Joliet"
    shtMasterJoliet.Range("F2") = shtInventory.Application.WorksheetFunction.IfError( _
        shtInventory.Application.WorksheetFunction.VLookup(shtMasterJoliet.Cells(2, 3), shtInventory.Columns("C:D"), _
            2, 0), 0)
    ' autofill the vlookup
    lastRow = shtInventory.Cells(Rows.Count, 3).End(xlUp).Row
    shtMasterJoliet.Range("F2").AutoFill Destination:=shtMasterJoliet.Range("F2:D" & lastRow)
            
    ' filter by modesto and perform vlookup
    shtInventory.ListObjects("Table2").Range.AutoFilter Field:=1, Criteria1:="Modesto"
    shtMasterModesto.Range("F2") = shtInventory.Application.WorksheetFunction.IfError( _
        shtInventory.Application.WorksheetFunction.VLookup(shtMasterModesto.Cells(2, 3), shtInventory.Columns("C:D"), _
            2, 0), 0)
    ' autofill the vlookup
    lastRow = shtInventory.Cells(Rows.Count, 3).End(xlUp).Row
    shtMasterModesto.Range("F2").AutoFill Destination:=shtMasterModesto.Range("F2:D" & lastRow)
    
    ' perform index match to put in ax numbers
    
'**************************************************************************************************************
    ' Step 3: Get the quantity from PO csv
        ' column O is ax number
        ' column P is description
        ' column R is quantity
        ' vlookup with ax number and return units, 0 if not found
        
    ' Step 4: get the quantity from TO
        ' column J is ax number
        ' column K is description
        ' column L is prod8
        ' column N is quantity
        ' vlookup with ax number, return quantity, 0 if not found
        
    ' Step 5: sum columns
        ' H + G + F
        
    ' Step 6: calculate difference
        ' I - E
        
    ' Save report
    ' set each variable to nothing
    shtMasterModesto = Nothing
    shtMasterJoliet = Nothing
    wkbInventory = Nothing
    shtInventory = Nothing
    wkbTransferOrder = Nothing
    shtTransferOrder = Nothing
    wkbPurchaseOrder = Nothing
    shtPurchaseOrder = Nothing
    wkbVbs = Nothing
    shtVbs = Nothing
    
End Sub

Attribute VB_Name = "NetInventoryMacro"
Public Sub NetInventory()
    'Clear the existing workbook
    Dim deleteSheet As Worksheet
    For Each deleteSheet In ActiveWorkbook.Worksheets
        If deleteSheet.Name = "Modesto" Then
            deleteSheet.Delete
        ElseIf deleteSheet.Name = "Joliet" Then
            deleteSheet.Name = "Sheet 1"
            deleteSheet.Cells.Clear
        ElseIf deleteSheet.Name = "Sheet 1" Then
            ' leave it be
        Else
            deleteSheet.Delete
        End If
    Next deleteSheet
        
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
    Sheets.Add After:=ActiveSheet
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
    Dim lastRow As Integer
    Dim invReportName As String
    Dim todayDate As Date, m As String, d As String, y As String
    todayDate = DateValue(Date)
    m = Month(todayDate)
    d = Day(todayDate)
    y = Year(todayDate)
    invReportName = m & "_" & d & "_" & y & "_InventoryReport.xlsx"
    
    Set wkbInventory = Workbooks.Open(path & invReportName)
    Set shtInventory = wkbInventory.Sheets("Daily Inventory")
    lastRow = shtMasterJoliet.Cells(Rows.Count, 3).End(xlUp).Row
    
    ' iterate through and get the correct quantity given the brewery is correct
    Dim prod8 As String, i As Integer, j As Integer, endRow As Integer
    Dim foundJoliet As Boolean, foundModesto As Boolean
    
    endRow = shtInventory.Cells(Rows.Count, 3).End(xlUp).Row
    
    ' iterate through each prod8 in new report
    For i = 2 To lastRow
        shtMasterJoliet.Cells(i, "A").Value = "Distribution Center 1"
        shtMasterModesto.Cells(i, "A").Value = "Distribution Center 1"
        prod8 = shtMasterJoliet.Cells(i, "C").Value
        foundJoliet = False
        foundModesto = False
        ' iterate through each row of inventory report and search for prod8
        For j = 2 To endRow
            'first check if the prod8's match
            If StrComp(shtInventory.Cells(j, "C").Value, prod8) = 0 Then
                'then check to see which brewery
                If StrComp(shtInventory.Cells(j, "A").Value, "Joliet") = 0 Then
                    ' if so, copy over the units and ax number to the new report
                    shtMasterJoliet.Cells(i, "F").Value = shtInventory.Cells(j, "D").Value
                    shtMasterJoliet.Cells(i, "B").Value = shtInventory.Cells(j, "B").Value
                    foundJoliet = True
                ElseIf StrComp(shtInventory.Cells(j, "A").Value, "Modesto") = 0 Then
                    ' if so, copy over the units
                    shtMasterModesto.Cells(i, "F").Value = shtInventory.Cells(j, "D").Value
                    foundModesto = True
                Else
                    'continue looping
                End If
            End If
            If foundModesto And foundJoliet Then j = endRow
        Next j
        ' if neither was found in the inventory report than mark units as 0 and ax as N/A
        If Not foundModesto Then
            shtMasterModesto.Cells(i, "F").Value = 0
        ElseIf Not foundJoliet Then
            shtMasterJoliet.Cells(i, "F").Value = 0
            shtMasterJoliet.Cells(i, "B").Value = "N/A"
        End If
    Next i

    ' copy ax numbers from joliet sheet to modesto sheet
    shtMasterJoliet.Range("B2:B" & lastRow).Copy Destination:=shtMasterModesto.Cells(2, "B")
    
    'close the inventory report
    wkbInventory.Close (False)
    
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
        shtMasterJoliet.Range("I2").Formula = "=$H2 + $G2 + $F2"
        shtMasterJoliet.Range("I2:I").FillDown
        
        shtMasterModesto.Range("I2").Formula = "=$H2 + $G2 + $F2"
        shtMasterModesto.Range("I2:I").FillDown
    ' Step 6: calculate difference
        ' I - E
        shtMasterJoliet.Range("J2").Formula = "=$I2-$E2"
        shtMasterJoliet.Range("J2:J").FillDown
        
        shtMasterModesto.Range("J2").Formula = "=$I2-$E2"
        shtMasterModesto.Range("J2:J").FillDown
    ' Save report
    ' set each variable to nothing

    
End Sub

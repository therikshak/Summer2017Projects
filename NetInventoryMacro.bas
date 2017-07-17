Attribute VB_Name = "NetInventoryMacro"
Public Sub NetInventory()
    ' put the following excel files into this folder
        ' daily inventory report
        ' transfer orders
        ' purchase orders'
        ' vbs copy paste
        
    ' open all of the excel files and store each one into a variable to reference
    
    ' in this report, columns will be:
    '   A   B    C       D           E                 F            G  H       I           J
    ' Plant|AX|Prod8|Description|Quantity(vbs)|quantity(inv report)|PO|TO|Total_projected|Diff
        ' total_projected = TO + PO + quantity(inv report)
        ' diff = total_projected - quanity(vbs)
    
    ' two sheets, one for modesto and one for joliet, repeat below for both sheets
    
    ' Step 1: Move data from vbs to this sheet
        ' leave off barrels
        ' range A2:D sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Step 2: get the inventory totals from daily report
        ' first filter the report to the correct brewery, then vlookup
        ' below goes into column F
        ' iferror(vlookup(B#,dailyinventory sheet( Range C2:D end of table), 2 (units), 0) , value if true = 0)
        ' after getting units, do another vlookup or index match to put ax numbers in
        
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
    ' close all files
    ' set each variable to nothing
End Sub

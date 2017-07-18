Attribute VB_Name = "DailyInventory"
Sub DailyInventory()
    Dim file As Variant, path As String
    Dim fileNames As New Collection
    Dim productInformationSheet As String
    Dim username As String
    
    username = (Environ$("Username"))
    path = "C:\Users\" & username & "\SharePoint\T\Projects\InventoryReports\"
    file = Dir(path)
    
    'get each file in the folder and store into collection
    Do While Len(file) > 0
        If InStr(1, file, "ProductInformation") > 0 Then
            productInformationSheet = file
        Else
            fileNames.Add file
        End If
        
        file = Dir
    Loop
    
    'create headings on sheet 1
    Range("A1") = "Brewery"
    Range("B1") = "AX #"
    Range("C1") = "Prod 8"
    Range("D1") = "Units"
    Range("E1") = "Production Date"
    Range("F1") = "Ship By Date"
    Range("G1") = "Alt SKU"
    Range("H1") = "Product Name"
    Range("I1") = "Product Description"

    
    Application.ScreenUpdating = False
    'loop through and open each file and run macro
    Dim shtMaster As Worksheet, shtGet As Worksheet, wkbGet As Workbook
    Dim i As Long, lengthMaster As Long, lengthGet As Long
    Set shtMaster = ActiveWorkbook.ActiveSheet
    
    For i = 1 To fileNames.Count
        'open the excel file
        Set wkbGet = Workbooks.Open(path & fileNames(i))
        Set shtGet = wkbGet.Sheets(1)
        
        'run correct macro and create table in workbook
        If InStr(1, fileNames(i), "AGED") > 0 Then
            cityInventory
        ElseIf InStr(1, fileNames(i), "Joliet") > 0 Then
            SaddlecreekInventory (False)
        ElseIf InStr(1, fileNames(i), "Modesto") > 0 Then
            SaddlecreekInventory (True)
        ElseIf InStr(1, fileNames(i), "New") > 0 Then
            newHolland
        ElseIf InStr(1, fileNames(i), "Strohs") > 0 Then
            brewDetroit
        ElseIf InStr(1, fileNames(i), "Lindner") > 0 Then
            Lindner
        ElseIf InStr(1, fileNames(i), "InventoryReport") > 0 Then
            'do nothing
        Else
            vermont
        End If
        'put data into master
        Set shtGet = wkbGet.Sheets(1)
        lengthGet = shtGet.Cells(Rows.Count, 1).End(xlUp).Row
        If i = 1 Then
            shtGet.Range("A1" & ":" & "H" & lengthGet).Copy Destination:=shtMaster.Cells(2, "A")
        Else
            lengthMaster = shtMaster.Cells(Rows.Count, 1).End(xlUp).Row + 1
            shtGet.Range("A1:H" & lengthGet).Copy Destination:=shtMaster.Cells(lengthMaster, "A")
        End If
        wkbGet.Close (False)
    Next i

'***********************************************************************************
'ADD SHIPBY DATE, AX# AND PROD8
    Dim wkbProdInfo As Workbook, shtProdInfoData As Worksheet, shtProdInfoDate As Worksheet
    productInformationSheet = "ProductInformation.xlsm"
    Set wkbProdInfo = Workbooks.Open(path & productInformationSheet)
    Set shtProdInfoData = wkbProdInfo.Sheets("Data")
    Set shtProdInfoDate = wkbProdInfo.Sheets("ShipBy")
    
    'shtMaster contains standard table
    'shtProdInfo(1) is the table of prod8 and ax
    'shtprodinfo(2) is table of ship by dates
    
    Dim Brewery As String
    Dim axNum As String, sku As String, prod8 As String, name As String
    Dim j As Long
    axNum = ""
    sku = ""
    prod8 = ""
    name = ""
    'Loop through shtMaster
    n = shtMaster.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To n
        'get brewery
        Brewery = shtMaster.Cells(i, 1).Text
        'if a city brewery
        If Brewery = "La Crosse, WI" Or Brewery = "Memphis, TN" Or Brewery = "Latrobe, PA" Then
            'get the ax number and sku from table
            axNum = shtMaster.Cells(i, 2)
            sku = shtMaster.Cells(i, 7)
            On Error Resume Next
            'perform a vlookup, ax number as lookup value, in the productInformation excel workbook
            'return the prod8, exact match
            shtMaster.Cells(i, 3) = shtProdInfoData.Application.WorksheetFunction.VLookup( _
            axNum, shtProdInfoData.Range("A1:C1000"), 3, 0)
            'if it does not find the ax number, then perform a vlookup with the sku instead
            If Err.Number <> 0 Then
                'try to find prod8 with sku
                shtMaster.Cells(i, 3) = shtProdInfoData.Application.WorksheetFunction.VLookup( _
                sku, shtProdInfoData.Range("B2:C1000"), 2, 0)
                'if it still does not find it, input N/A
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 3) = "N/A"
                End If
            End If
            On Error GoTo 0
        'if a saddlecreek brewery
        ElseIf Brewery = "Joliet" Or Brewery = "Modesto" Then
            'get the prod8 from the table
            prod8 = shtMaster.Cells(i, 3)
            On Error Resume Next
            'perform an index match with the prod8 as the lookup value
            'return the ax number
            With shtProdInfoData.Application.WorksheetFunction
                shtMaster.Cells(i, 2) = _
                .Index(shtProdInfoData.Range("A2:A1000"), _
                .Match(prod8, shtProdInfoData.Range("C2:C1000"), 0))
            End With
            If Err.Number <> 0 Then
                shtMaster.Cells(i, 2) = "N/A"
            End If
            On Error GoTo 0
        'if new holland
        ElseIf Brewery = "New Holland" Then
            'get the name of the product from the table
            name = shtMaster.Cells(i, 8)
            On Error Resume Next
            'perform an index match with the name to get the ax number
            'then perform a second index match to get the prod8
            With shtProdInfoData.Application.WorksheetFunction
                shtMaster.Cells(i, 2) = _
                .Index(shtProdInfoData.Range("A2:A1000"), _
                .Match(name, shtProdInfoData.Range("F2:F1000"), 0))
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 2) = "N/A"
                End If
                shtMaster.Cells(i, 3) = _
                .Index(shtProdInfoData.Range("C2:C1000"), _
                .Match(name, shtProdInfoData.Range("F2:F1000"), 0))
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 3) = "N/A"
                End If
            End With
        ElseIf Brewery = "Brew Detroit" Then
            name = shtMaster.Cells(i, 8)
            On Error Resume Next
            'perform an index match with the name to get the ax number
            'then perform a second index match to get the prod8
            With shtProdInfoData.Application.WorksheetFunction
                shtMaster.Cells(i, 2) = _
                .Index(shtProdInfoData.Range("A2:A1000"), _
                .Match(name, shtProdInfoData.Range("F2:F1000"), 0))
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 2) = "N/A"
                End If
                shtMaster.Cells(i, 3) = _
                .Index(shtProdInfoData.Range("C2:C1000"), _
                .Match(name, shtProdInfoData.Range("F2:F1000"), 0))
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 3) = "N/A"
                End If
            End With
        ElseIf Brewery = "Lindner" Then
            axNum = shtMaster.Cells(i, 2)
            On Error Resume Next
            'perform an index match to get the prod8
            With shtProdInfoData.Application.WorksheetFunction
                shtMaster.Cells(i, 3) = _
                .Index(shtProdInfoData.Range("C2:C1000"), _
                .Match(axNum, shtProdInfoData.Range("A2:A1000"), 0))
                If Err.Number <> 0 Then
                    shtMaster.Cells(i, 3) = "N/A"
                End If
            End With
        Else 'vermont
            'get the prod8 from the table
            prod8 = shtMaster.Cells(i, 3)
            'perform an index match with the prod8 to get the ax number
            With shtProdInfoData.Application.WorksheetFunction
                shtMaster.Cells(i, 2) = _
                .Index(shtProdInfoData.Range("A2:A1000"), _
                .Match(prod8, shtProdInfoData.Range("C2:C1000"), 0))
            End With
            GoTo vermont
        End If
        'get the shipBy date
        name = shtMaster.Cells(i, 8).Value 'name of product
        'get numbers of cells in the shipby date table
        M = shtProdInfoDate.Cells(Rows.Count, 1).End(xlUp).Row
        For j = 2 To M
            If InStr(1, name, shtProdInfoDate.Cells(j, 1).Text) > 0 Then
                shtMaster.Cells(i, 6).Value = shtMaster.Cells(i, 5).Value + _
                shtProdInfoDate.Cells(j, 2).Value
                GoTo foundShipBy
            End If
        Next j
vermont:
        If shtMaster.Cells(i, 5).Value = "NO DATA" Then
            shtMaster.Cells(i, 6).Value = "NO DATA"
        Else
            shtMaster.Cells(i, 6).Value = "NO DATA"
        End If
foundShipBy:
    Next i
    
    'ADD PRODUCT DESCRIPTION
    For i = 2 To n
        axNum = shtMaster.Cells(i, 2).Value
        prod8 = shtMaster.Cells(i, 3).Value
        If axNum = "N/A" Then
            If prod8 = "N/A" Then
                GoTo default
            Else
                GoTo prod8Search
            End If
        End If
        On Error Resume Next
            'perform a vlookup, ax number as lookup value, in the productInformation excel workbook
            'return the description, exact match
            shtMaster.Cells(i, 9) = shtProdInfoData.Application.WorksheetFunction.VLookup( _
            axNum, shtProdInfoData.Range("A2:D1000"), 4, 0)
            'if it does not find the ax number, then perform a vlookup with the prod8 instead
            If Err.Number <> 0 Then
prod8Search:
                shtMaster.Cells(i, 9) = shtProdInfoData.Application.WorksheetFunction.VLookup( _
                prod8, shtProdInfoData.Range("C2:D1000"), 2, 0)
                'if it still does not find it, use what is in column H
                If Err.Number <> 0 Then
default:
                    shtMaster.Cells(i, 9) = shtMaster.Cells(i, 8)
                End If
            End If
        'convert the ax to a number
        shtMaster.Cells(i, 2).Value = Val(shtMaster.Cells(i, 2).Value)
    Next i
    On Error GoTo 0
    
    'close the prodinfo workbook
    wkbProdInfo.Close (False)
    Set wkbProdInfo = Nothing
    Set wkbGet = Nothing
    
    'create table with dates
    DailyInventoryTableDates
    
    'Sort sheet by axnum for efficiently creating new sheet
    With shtMaster.ListObjects(1).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Table1[AX '#]"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("Table1[Prod 8]"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("Table1[Brewery]"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'create minimal table without dates
    DailyInventoryNoDates

End Sub

'**************************************************************************
'SADDLECREEK

'if boolean modesto is true, then 1 year needs to be subtracted from the dates
Private Sub SaddlecreekInventory(ByVal Modesto As Boolean)
    'Arrays to be filled with data
    Dim itemArr() As String, itemArrSize As Integer
    Dim skuArr() As String
    Dim itemStartCellArr() As Integer
    Dim itemEndCellArr() As Integer
    Dim numProdDatesArr() As Integer, numProdDatesArrSize As Integer
    
'**************************************************************************
'SetUp the excel sheet
    'Unmerge all cells to make them easier to work with
    ActiveSheet.Cells.UnMerge
    Dim version As Integer
    version = 0
    If Cells(6, "K") = "Lot04" Then
        version = 2
    Else
        version = 1
    End If
    'delete first 6 rows for formatting
    Rows("1:6").Delete
'**************************************************************************
'Get the item names and SKU
    'Loop through Column B to get each item
    'Test to make sure data is in the cell before copying
    'Store #rows between each to know how many rows between items
    Dim i As Long, n As Long
    n = Cells(Rows.Count, 1).End(xlUp).Row
    itemArrSize = 0
    skuArrSize = 0
    
    For i = 1 To n
        If Not IsEmpty(Cells(i, "B").Value) Then
            'increment array size
            itemArrSize = itemArrSize + 1
            'reallocate the arrays
            ReDim Preserve itemArr(itemArrSize)
            ReDim Preserve itemStartCellArr(itemArrSize)
            ReDim Preserve itemEndCellArr(itemArrSize)
            ReDim Preserve skuArr(itemArrSize)

            'add name of item to item array and start cell of that item
            itemArr(itemArrSize - 1) = Cells(i, "B").Value
            skuArr(itemArrSize - 1) = Cells(i, "A").Value
            itemStartCellArr(itemArrSize - 1) = i
            If itemArrSize > 1 Then
                itemEndCellArr(itemArrSize - 2) = (i - 1)
            End If
        End If
    Next i
    itemEndCellArr(itemArrSize - 1) = n
'**************************************************************************
'Get Production Dates
    'production date arrays
    Dim prodDates As New Collection
    Dim thisProdDateArr() As Variant
    
    'item index keeps track of which item the date is for
    'thisProdDateNum keeps track of the number of production dates for
    'each individual item
    Dim itemIndex As Long, j As Long, thisProdDateNum As Long
    itemIndex = 0
    
    'Loop through Column L for production dates
    Do While itemIndex < itemArrSize
        itemIndex = itemIndex + 1
        thisProdDateNum = 0
        'reset thisProdDateNum array
        ReDim thisProdDateArr(0)
        For j = itemStartCellArr(itemIndex - 1) To itemEndCellArr(itemIndex - 1)
            If version = 1 Then
                If Cells(j, "L").Value <> "" Then
                    thisProdDateNum = thisProdDateNum + 1
                    ReDim Preserve thisProdDateArr(thisProdDateNum - 1)
                    If Modesto Then
                        thisProdDateArr(thisProdDateNum - 1) = DateAdd("yyyy", -1, Int(Cells(j, "L").Value))
                    Else
                        thisProdDateArr(thisProdDateNum - 1) = Int(Cells(j, "L").Value)
                    End If
                End If
            Else
                If Cells(j, "K").Value <> "" Then
                    thisProdDateNum = thisProdDateNum + 1
                    ReDim Preserve thisProdDateArr(thisProdDateNum - 1)
                    If Modesto Then
                        thisProdDateArr(thisProdDateNum - 1) = DateAdd("yyyy", -1, Int(Cells(j, "K").Value))
                    Else
                        thisProdDateArr(thisProdDateNum - 1) = Int(Cells(j, "K").Value)
                    End If
                End If
            End If
        Next j
        prodDates.Add (thisProdDateArr)
    Loop
'**************************************************************************
'Get Inventory Totals
    Dim totalByDate As New Collection
    Dim thisTotalArr() As Variant
    
    Dim thisTotalNum As Long
    itemIndex = 0
    
    'Loop through Column S to get # units for each production date
    Do While itemIndex < itemArrSize
        itemIndex = itemIndex + 1
        thisTotalNum = 0
        'reset this item's totals array
        ReDim thisTotalArr(0)
        For j = itemStartCellArr(itemIndex - 1) To itemEndCellArr(itemIndex - 1)
            If version = 1 Then
                'if the font is bold
                If Cells(j, "S").Font.Bold = True Then
                    'but it is not the grand total for that item
                    If j <> itemEndCellArr(itemIndex - 1) Then
                        'add the total to this item's array
                        thisTotalNum = thisTotalNum + 1
                        ReDim Preserve thisTotalArr(thisTotalNum - 1)
                        thisTotalArr(thisTotalNum - 1) = Cells(j, "S").Value
                    End If
                End If
            Else
                'if the font is bold
                If Cells(j, "R").Font.Bold = True Then
                    'but it is not the grand total for that item
                    If j <> itemEndCellArr(itemIndex - 1) Then
                        'add the total to this item's array
                        thisTotalNum = thisTotalNum + 1
                        ReDim Preserve thisTotalArr(thisTotalNum - 1)
                        thisTotalArr(thisTotalNum - 1) = Cells(j, "R").Value
                    End If
                End If
            End If
        Next j
        totalByDate.Add (thisTotalArr)
    Loop
'**************************************************************************
'Combine production dates and inventory totals
    Dim finalInventory As New Collection
    Dim itemInv As Scripting.Dictionary
    
    'i keeps track of which item
    i = 1
    'Add Dates and number produced at that date to final Inventory
    For Each collectionItem In prodDates
        'j keeps track of array index of each item
        j = 0
        Set itemInv = New Scripting.Dictionary
        For Each element In collectionItem
            If itemInv.Exists(element) Then
                itemInv(element) = itemInv(element) + totalByDate(i)(j)
                j = j + 1
            Else
                itemInv.Add element, totalByDate(i)(j)
                j = j + 1
            End If
        Next element
    'add this item's dictionary to theF collection
    finalInventory.Add itemInv
    i = i + 1
    Next collectionItem
'**************************************************************************
'Create Standard Format Excel Table With the Data
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    'Get production dates and count for each item and print to sheet
    'each iteration is a new item
    
    i = 0 'for array index of item name
    j = 1 'to keep track of row
    For Each collectionItem In finalInventory
    
        For Each Key In collectionItem.Keys()
            If Modesto Then
                Cells(j, 1).Value = "Modesto"
            Else
                Cells(j, 1).Value = "Joliet"
            End If
            Cells(j, 3).Value = skuArr(i)
            Cells(j, 7).Value = skuArr(i)
            Cells(j, 8).Value = itemArr(i)
            Cells(j, 5).Value = Key
            Cells(j, 4).Value = collectionItem(Key)
            j = j + 1
        Next Key
        i = i + 1
    Next collectionItem
End Sub

'**************************************************************************
'LINDNER
Private Sub Lindner()
    Dim bName As String
    Dim axNum As New Collection
    Dim prodName As New Collection
    Dim quantity As New Collection
    
    bName = "Lindner"
    pDate = "NO DATA"
    'delete first 8 rows
    Rows("1:8").Delete Shift:=xlUp
    'get number of rows in table
    n = Cells(Rows.Count, "E").End(xlUp).Row
    'loop through table and extract information
    For i = 2 To n
        axNum.Add Cells(i, "A").Value
        prodName.Add Cells(i, "B").Value
        quantity.Add Cells(i, "E").Value
    Next i
    
    'output to standard table
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    For j = 1 To prodName.Count
        Cells(j, 1).Value = bName
        Cells(j, 8).Value = prodName(j)
        Cells(j, 5).Value = pDate
        Cells(j, 4).Value = quantity(j)
        Cells(j, 2).Value = axNum(j)
    Next j
End Sub

'**************************************************************************
'CITY
Private Sub cityInventory()
    'unmerge all cells and unwrap text
    ActiveSheet.Cells.UnMerge
    ActiveSheet.Cells.WrapText = False
    'get brewery name
    Dim bName As String
    bName = Cells(2, 1).Value
'**************************************************************************
'Get Product Names
    'Create dictionary for product names
    Dim prodNames As New Scripting.Dictionary
    Dim axNum As New Collection
    Dim cityNum As New Collection
    Dim i As Long, n As Long, j As Long
    n = Cells(Rows.Count, "E").End(xlUp).Row
    For i = 7 To n
        If prodNames.Exists(Cells(i, "E").Value) Then
            'increment count
            prodNames(Cells(i, "E").Value) = prodNames(Cells(i, "E").Value) + 1
        Else
            'add the product as the key
            prodNames.Add Cells(i, "E").Value, 1
            axNum.Add Cells(i, "D")
            cityNum.Add Cells(i, "C")
        End If
    Next i
'**************************************************************************
'Get Production Dates and Quantities
    Dim inventory As New Collection
    Dim dateNum As Scripting.Dictionary
    Dim prodCount As Integer
    
    i = 7
    For Each Key In prodNames.Keys()
        Set dateNum = New Scripting.Dictionary
        'number of productions for each item
        prodCount = prodNames(Key)
        For j = 1 To prodCount
            'if date is added for this product, increment amount
            If dateNum.Exists(Cells(i, "M").Value) Then
                dateNum(Cells(i, "M").Value) = dateNum(Cells(i, "M").Value) + Cells(i, "J").Value
                i = i + 1
            'otherwise add date and count
            Else
                dateNum.Add Cells(i, "M").Value, Cells(i, "J").Value
                i = i + 1
            End If
        Next j
        inventory.Add dateNum
    Next Key
'**************************************************************************
'Output information
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    j = 1 'to keep track of row to output to
    k = 1 'to keep track of item
    
    'writes out every product name and the brewery
    For Each Key In prodNames.Keys()
        For i = 1 To inventory(k).Count
            Cells(j, 8).Value = Key
            Cells(j, 1).Value = bName
            Cells(j, 7).Value = cityNum(k)
            Cells(j, 2).Value = axNum(k)
            j = j + 1
        Next i
        k = k + 1
    Next Key
    
    'writes out every production date and # of units
    j = 1
    For Each collectionItem In inventory
        For Each Key In collectionItem.Keys()
            Cells(j, 5).Value = Key
            Cells(j, 4).Value = collectionItem(Key)
            j = j + 1
        Next Key
    Next collectionItem
    
    'Get rid of extra characters in axNums and city prod nums
    Columns("G").Replace _
        What:="F", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Columns("B").Replace _
        What:="a", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Columns("B").Replace _
        What:="b", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Columns("B").Replace _
        What:=".", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
End Sub

'******************************************************************************
'BREW DETROIT

Private Sub brewDetroit()
    'get brewery name
    Dim bName As String
    bName = Trim(Cells(4, "J").Value)
    
'Get product name
    Dim prodName As New Collection
    Dim prod As String
    Dim i As Long
    i = 6
    Do Until Cells(i, 1).Value = "Totals"
        prod = Cells(i, 1).Value & Cells(i, 2).Value & " " & Cells(i, 3).Value
        prodName.Add prod
        i = i + 1
    Loop
    
'Set N/A for date
    Dim prodDate As String
    prodDate = "NO DATA"

'Set # of Units for each product
    i = 6
    Dim numUnits As New Collection
    Do Until Cells(i, "J").Font.Bold = True
        numUnits.Add Cells(i, "J").Value
        i = i + 1
    Loop
'**************************************************************************
'Output data to a Table
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    For i = 1 To prodName.Count
        Cells(i, 1).Value = bName
        Cells(i, 8).Value = prodName(i)
        Cells(i, 5).Value = prodDate
        Cells(i, 4).Value = numUnits(i)
    Next i
End Sub

'**************************************************************************
'VERMONT

Private Sub vermont()
    'unmerge all cells and unwrap text
    ActiveSheet.Cells.UnMerge
    'get brewery name
    Dim bName As String, pDate As String
    pDate = "NO DATA"
    bName = "Vermont Cider"
    
'Get product names
    Dim prodNames As New Collection
    Dim prod8 As New Collection
    Dim r As Range
    Set r = ActiveSheet.Range("A1:Z400")
    
    n = Cells(Rows.Count, "H").End(xlUp).Row
    Dim i As Long
    'delete empty rows
    For i = 9 To n + 1
        If WorksheetFunction.CountA(r.Rows(i)) = 0 Then
            r.Rows(i).Delete
        End If
    Next i
    
    n = Cells(Rows.Count, "H").End(xlUp).Row
    For i = 9 To n + 1
        prodNames.Add Cells(i, "H").Value
        prod8.Add Cells(i, "F").Value
    Next i
    
'Get number of units
    Dim numUnits As New Collection
    For i = 9 To n
        numUnits.Add Cells(i, "J").Value
    Next i
'**************************************************************************
'Output Data to standard table
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    Dim j As Integer
    j = 1
    For i = 2 To prodNames.Count
        Cells(i - 1, 1).Value = bName
        Cells(i - 1, 3).Value = prod8(j)
        Cells(i - 1, 8).Value = prodNames(j)
        Cells(i - 1, 5).Value = pDate
        Cells(i - 1, 4).Value = numUnits(j)
        j = j + 1
    Next i
End Sub

'**************************************************************************
' NEW HOLLAND

Private Sub newHolland()
    Dim bName As String
    bName = "New Holland"
'Get product names, production dates, and units
    Dim prodNames As New Collection
    Dim prodDates As New Collection
    Dim units As New Collection
    Dim i As Long, n As Long
    
    n = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To n
        prodNames.Add Cells(i, "E").Value
        prodDates.Add Cells(i, "B").Value
        units.Add Cells(i, "F").Value
    Next i
'**************************************************************************
'Output information to formatted table
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    ws.name = "Table"
    Dim j As Integer
    j = 1
    For i = 2 To prodNames.Count + 1
        Cells(i - 1, 1).Value = bName
        Cells(i - 1, 8).Value = prodNames(j)
        Cells(i - 1, 5).Value = prodDates(j)
        Cells(i - 1, 4).Value = units(j)
        j = j + 1
    Next i
End Sub

'**************************************************************************
'CREATE MINIMAL TABLE

Private Sub DailyInventoryNoDates()
    'newSheet is for table with no dates
    'dataSheet is the existing table with all info
    Dim newSheet As Worksheet, dataSheet As Worksheet
    'add the new sheet before the existing one
    Sheets.Add Before:=ActiveSheet
    ActiveSheet.name = "Daily Inventory"
    Set newSheet = Worksheets("Daily Inventory")
    Set dataSheet = Worksheets(2)
    dataSheet.name = "Daily Inventory With Dates"
    
    'HEADERS FOR NEW TABLE
    newSheet.Range("A1") = "Brewery"
    newSheet.Range("B1") = "AX #"
    newSheet.Range("C1") = "Prod 8"
    newSheet.Range("D1") = "Units"
    newSheet.Range("E1") = "Product Description"
    
    'Create array for data
    Dim i As Long, lastRow As Long
    
    Dim bName As String, axNum As String, prod8 As String
    Dim numUnits As Integer, prodDescr As String, inc As Double, orig As Double
    
    Dim data As New Collection ''''''''''''''''''''
    Dim arr(4)
    
    'add first row of data before looping
    arr(0) = dataSheet.Cells(2, 1).Value
    arr(1) = dataSheet.Cells(2, 2).Value
    'prod8
    arr(2) = dataSheet.Cells(2, 3).Value
    'units
    arr(3) = dataSheet.Cells(2, 4).Value
    'description
    arr(4) = dataSheet.Cells(2, 9).Value
    data.Add (arr)
    lastRow = dataSheet.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 3 To lastRow
        'get brewery name
        bName = dataSheet.Cells(i, 1).Value
        'if brewery on this iteration is equal to last element added to collection
        If bName = data(data.Count)(0) Then
            'get axnum
            axNum = dataSheet.Cells(i, 2).Value
            If axNum = 0 Then
                prod8 = dataSheet.Cells(i, 3).Value
                If prod8 = "N/A" Then
                    prodDescr = dataSheet.Cells(i, 9).Value
                    If prodDescr = data(data.Count)(4) Then
                        GoTo increment
                    Else
                        GoTo addNew
                    End If
                ElseIf prod8 = data(data.Count)(2) Then
                    GoTo increment
                Else
                    GoTo addNew
                End If
            'if axnum = last axnum added
            ElseIf axNum = data(data.Count)(1) Then
                GoTo increment
            'otherwise add new entry
            Else
                GoTo addNew
            End If
        Else
            GoTo addNew
        End If
increment:
      inc = dataSheet.Cells(i, 4).Value
      orig = data(data.Count)(3)
      arr(3) = orig + inc
      data.Remove (data.Count)
      data.Add (arr)
      GoTo nextIt
addNew:
    arr(0) = bName
    'ax
    arr(1) = dataSheet.Cells(i, 2).Value
    'prod8
    arr(2) = dataSheet.Cells(i, 3).Value
    'units
    arr(3) = dataSheet.Cells(i, 4).Value
    'description
    arr(4) = dataSheet.Cells(i, 9).Value
    data.Add (arr)
nextIt:
    Next i

    'OUTPUT TABLE DATA
    For i = 1 To data.Count
        newSheet.Cells(i + 1, 1) = data(i)(0)
        newSheet.Cells(i + 1, 2) = data(i)(1)
        newSheet.Cells(i + 1, 3) = data(i)(2)
        newSheet.Cells(i + 1, 4) = data(i)(3)
        newSheet.Cells(i + 1, 5) = data(i)(4)
    Next i
    
    'TURN INTO TABLE
    n = newSheet.Cells(Rows.Count, 1).End(xlUp).Row
    newSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$E$" & n), , xlYes).name = "Table2"
    With Columns("A:D")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .EntireColumn.AutoFit
    End With
    With Columns("E:E")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 40
    End With

    'change row heights
    ActiveSheet.Range("A2:A" & newSheet.Rows.Count).RowHeight = 30
    
    'SLICERS
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "AX #"). _
        Slicers.Add ActiveSheet, , "AX # 1", "AX #", 210, 470, 120, 200
    
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "Prod 8"). _
        Slicers.Add ActiveSheet, , "Prod 8 1", "Prod 8", 210, 590, 120, 200
        
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "Brewery"). _
        Slicers.Add ActiveSheet, , "Brewery 1", "Brewery", 210, 710, 120, 200
    
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), _
        "Product Description").Slicers.Add ActiveSheet, , "Product Description 1", _
        "Product Description", 5, 470, 360, 200
End Sub

'**************************************************************************
'CREATE TABLE WITH DATES

Private Sub DailyInventoryTableDates()
    n = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To n
        'if sku cell is empty add N/A
        If IsEmpty(Cells(i, 7).Value) = True Then
            Cells(i, 7).Value = "N/A"
        End If
    Next i
        
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$I$" & n), , xlYes).name = "Table1"
    Columns("I:I").EntireColumn.AutoFit
    With Columns("A:F")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .EntireColumn.AutoFit
    End With
    With Columns("I:I")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 40
    End With
    'hide auxiliary columns
        'hide date and auxiliary columns
        'Columns("E:H").EntireColumn.Hidden = True
    Columns("G:H").EntireColumn.Hidden = True
    
    'change row heights
    ActiveSheet.Range("A2:A" & ActiveSheet.Rows.Count).RowHeight = 30
    
    'SLICERS
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table1"), "AX #"). _
        Slicers.Add ActiveSheet, , "AX #", "AX #", 210, 660, 120, 200
        
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table1"), "Prod 8"). _
        Slicers.Add ActiveSheet, , "Prod 8", "Prod 8", 210, 780, 120, 200
        
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table1"), "Brewery"). _
        Slicers.Add ActiveSheet, , "Brewery", "Brewery", 210, 900, 120, 200
        
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table1"), _
        "Product Description").Slicers.Add ActiveSheet, , "Product Description", "Product Description", _
        5, 660, 360, 200
    
    'HIDE COLUMNS
        'ActiveSheet.Buttons.Add(510, 5, 60, 40).OnAction = "HideDate"
        'ActiveSheet.Buttons.Add(570, 5, 60, 40).OnAction = "ShowDate"
        
        'With ActiveSheet.Shapes("Button 1")
            '.Placement = xlFreeFloating
            '.TextFrame.Characters.Text = "Hide Dates"
        'End With
        'With ActiveSheet.Shapes("Button 2")
            '.Placement = xlFreeFloating
            '.TextFrame.Characters.Text = "Show Dates"
        'End With

End Sub

Private Sub HideDate()
        Columns("E:F").EntireColumn.Hidden = True
        ActiveSheet.Shapes("Product Description").IncrementLeft -180
        ActiveSheet.Shapes("AX #").IncrementLeft -180
        ActiveSheet.Shapes("Prod 8").IncrementLeft -180
        ActiveSheet.Shapes("Button 2").IncrementLeft -180
        ActiveSheet.Shapes("Button 1").IncrementLeft -180
End Sub

Private Sub ShowDate()
        Columns("E:F").EntireColumn.Hidden = False
        ActiveSheet.Shapes("Product Description").IncrementLeft 180
        ActiveSheet.Shapes("AX #").IncrementLeft 180
        ActiveSheet.Shapes("Prod 8").IncrementLeft 180
        ActiveSheet.Shapes("Button 2").IncrementLeft 180
        ActiveSheet.Shapes("Button 1").IncrementLeft 180
End Sub


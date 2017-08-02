Attribute VB_Name = "RunMillerReport"
Sub MillerCoorsOrderSummary()
'
' MillerCoorsOrderSummaryPivotTable Macro
' Creates a Pivot Table to view orders from selected product and breweries
    
    Dim pSheet As Worksheet
    Dim dSheet As Worksheet
    Dim wkbFinal As Workbook
    Dim pRange As Range
    Dim lastRow As Long
    Dim lastcol As Long
    
    Set pSheet = Worksheets("Pivot Table")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' delete the old data sheet if it is there
    On Error Resume Next
        ThisWorkbook.Worksheets("final").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    '  copy the data from the csv to this workbook
    Set wkbFinal = Workbooks.Open(ThisWorkbook.path & "\final.csv")
    wkbFinal.Sheets(1).Copy After:=pSheet
    wkbFinal.Close (False)
    Set dSheet = Worksheets("final")
    
    lastRow = dSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastcol = 9
    Set pRange = dSheet.Cells(1, 1).Resize(lastRow, lastcol)
    
    Dim i As Long
    For i = 2 To lastRow
        If (InStr(1, dSheet.Cells(i, 7).Value, "Order") > 0) Then
            dSheet.Cells(i, 9).Value = dSheet.Cells(i, 9).Value * -1
        End If
    Next i
              
    pSheet.PivotTables("MillerCoorsPivot").ChangePivotCache _
        ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=pRange)
    pSheet.PivotTables("MillerCoorsPivot").RefreshTable
    pSheet.Columns("B:K").ColumnWidth = 15
    pSheet.Activate
End Sub


